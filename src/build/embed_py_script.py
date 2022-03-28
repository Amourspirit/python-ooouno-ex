# coding: utf-8
import os
import zipfile
import shutil
from typing import Union, List, Optional
from pathlib import Path
from config import AppConfig
from .tmp_dir import tmpdir
from .copy_resource import CopyResource
from ..utils import util


class EmbedScriptPy:
    def __init__(
        self,
        src: Union[str, Path, List[str]],
        doc_path: Union[str, Path, List[str]],
        config: Optional[AppConfig] = None,
    ) -> None:
        self._src = util.get_path(src, ensure_absolute=True)
        self._doc_path = util.get_path(doc_path, ensure_absolute=True)
        self._config: AppConfig = config
        if self._config is None:
            self._config = util.get_app_cfg()

    def embed(self) -> None:
        def unzip(source: Path, dest: Path) -> None:
            with zipfile.ZipFile(str(source), "r") as zip_ref:
                zip_ref.extractall(str(dest))

        def zipdir(unzipped_path: Path, ziph: zipfile.ZipFile):
            for root, dirs, files in os.walk(unzipped_path):
                if len(files) == 0:
                    # it is a bit tricky adding empyty dir to zip.
                    # this block add the empty dirs uising ZipInfo
                    p_root = Path(root)
                    rel = p_root.relative_to(unzipped_path.parent)
                    zfi = zipfile.ZipInfo.from_file(p_root, rel)
                    ziph.writestr(zfi, "")
                    continue
                for file in files:
                    ziph.write(
                        os.path.join(root, file),
                        os.path.relpath(
                            os.path.join(root, file), os.path.join(unzipped_path, "..")
                        ),
                    )

        def zip_dir(unzipped_path: Path, dst_zip: Path) -> None:
            # https://ask.libreoffice.org/t/libreoffice-6-0-unzip-zip-open-document-file/40028
            # mimefile must be first file in zip.
            files = [p for p in unzipped_path.iterdir() if p.is_file()]
            subdirs = [p for p in unzipped_path.iterdir() if p.is_dir()]
            with zipfile.ZipFile(dst_zip, "w", zipfile.ZIP_DEFLATED) as zipf:
                for file in files:

                    pfile = file
                    zipf.write(pfile, pfile.relative_to(unzipped_path))
                for d in subdirs:
                    zipdir(unzipped_path=d, ziph=zipf)

        def zip_append(zip_path: Path, script_path: Path, tmpdir: Path) -> None:
            rel = script_path.relative_to(tmpdir)
            with zipfile.ZipFile(zip_path, "a") as zipf:
                zipf.write(script_path, rel)

        def get_odt_dest_path(tmp_dir: Path) -> Path:
            src = self._doc_path
            dst = Path(tmp_dir, f"{src.stem}.zip")
            return dst

        def embed_script(script_src: Path, zip_extract_dst: Path) -> None:
            p_script = zip_extract_dst / "Scripts" / "python"
            dst = Path(p_script, script_src.name)
            util.mkdirp(p_script)
            shutil.copy2(script_src, dst)

        def copy_script(script_src: Path, tmpdir: Path) -> None:
            p_script = tmpdir / "Scripts" / "python"
            dst = Path(p_script, script_src.name)
            util.mkdirp(p_script)
            shutil.copy2(script_src, dst)

        self._validate()
        with tmpdir() as temp_dir:
            p_tmp = Path(temp_dir)

            cp = CopyResource(
                src=self._doc_path,
                dst=get_odt_dest_path(p_tmp),
                dst_is_file=True,
                clear_prev=False,
                config=self._config,
            )
            cp.copy()
            copy_script(script_src=self._src, tmpdir=p_tmp)
            zip_path = Path(p_tmp, f"{cp.dst_path.stem}.zip")
            script_path = cp.dst_path.parent / "Scripts" / "python" / self._src.name
            zip_append(zip_path=zip_path, script_path=script_path, tmpdir=p_tmp)
            # print("done")
            # return
            zip_path = cp.dst_path
            zip_extract_dst = Path(p_tmp, zip_path.stem)
            unzip(source=zip_path, dest=zip_extract_dst)
            if zip_path.exists():
                os.remove(zip_path)
            embed_script(script_src=self._src, zip_extract_dst=zip_extract_dst)
            zip_dir(unzipped_path=zip_extract_dst, dst_zip=cp.dst_path)
            return

    def _validate(self) -> None:
        if self._src.exists() is False:
            raise FileNotFoundError(str(self._src))
