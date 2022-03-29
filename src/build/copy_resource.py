# coding: utf-8
import os
from distutils.dir_util import copy_tree
import shutil
from pathlib import Path
from typing import Union, List, Optional
from config import AppConfig
from ..utils import util


class CopyResource:
    """
    Copies a resource from the source or into destination dir.
    Resource can be a file or a dir.
    """

    def __init__(
        self,
        src: Union[str, Path, List[str]],
        dst: Union[str, Path, List[str], None],
        dst_is_file: bool = False,
        clear_prev: bool = True,
        src_is_res: bool = True,
        config: Optional[AppConfig] = None,
    ) -> None:
        """
        Constructor

        Arguments:
            src (str, Path, List[str]): path to the resource to copy. This can be a path to a file or dir.
            dst (str, Path, List[str]): destination path to copy resource.
            dst_is_file (bool, Optional): Specifies if ``dst`` is a file or a dir. Default ``False``
            clear_previous (bool, optional): If ``True`` previous files/dir will be cleared before resources are copied. Default: ``True``
            src_is_res (bool, Optional): Specifies if ``src`` is in the resource. Default ``True``
            config (AppConfig, Optional): App config. Default ``None``
        Raises:
            FileNotFoundError: If resource path is not found.
        """
        self._config: AppConfig = config
        if self._config is None:
            self._config = util.get_app_cfg()
        self._clear_previous = clear_prev
        _src = util.get_path(src)
        self._dst_is_file = dst_is_file
        if dst is None:
            _dst = None
        else:
            _dst = util.get_path(dst, ensure_absolute=True)

        if src_is_res:
            if _src.is_absolute():
                self._src = _src
            else:
                res_path = util.get_path(self._config.app_res_dir, ensure_absolute=True)
                self._src = Path(res_path, _src)
        else:
            self._src = util.get_path(_src, ensure_absolute=True)
        if not self._src.exists():
            raise FileNotFoundError(
                f"{self.__class__.__name__} unable to find resource path: '{self._src}'"
            )
        if _dst is None:
            self._dst = None
        else:
            if self._dst_is_file:
                self._dst = _dst
            else:
                self._dst = Path(_dst, self._src.name)

    def copy(self) -> None:
        """
        Copies source to dest.
        """
        if self._dst is None:
            raise ValueError("Destination not set")
        if self._src.is_dir():
            self._copy_dir()
        elif self._src.is_file():
            self._copy_file()

    def _copy_dir(self) -> None:
        if self._clear_previous == True and self._dst.exists():
            shutil.rmtree(self._dst)
        util.mkdirp(self._dst)
        copy_tree(str(self._src), str(self._dst))

    def _copy_file(self) -> None:
        p_dir = self._dst.parent
        util.mkdirp(p_dir)
        self._remove_files_in_dir(p_dir)
        shutil.copy2(self._src, self._dst)

    def _remove_files_in_dir(self, src_dir: Path) -> None:
        if self._clear_previous == False:
            return None
        if src_dir.exists() is False:
            return None
        for f in os.listdir(src_dir):
            f_path = Path(src_dir, f)
            if f_path.is_file():
                os.remove(f_path)

    @property
    def dst_path(self) -> Union[Path, None]:
        """
        Gets the path to the dest directory/file.
        """
        return self._dst

    @property
    def src_path(self) -> Path:
        """
        Gets the path to the dest directory/file.
        """
        return self._src

    @property
    def clear_previous(self) -> bool:
        """Gets if the previous files will be cleared"""
        return self._clear_previous
