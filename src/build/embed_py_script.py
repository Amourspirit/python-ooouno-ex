# coding: utf-8
import os
import zipfile
import shutil
from typing import Union, List, Optional
from pathlib import Path
from config import AppConfig
from .tmp_dir import tmpdir
from .copy_resource import CopyResource
from .manifest_script import ManifestScript
from ..utils import util
from ..models.example.model_example import ModelExample as Model
from . import build_util

class EmbedScriptPy:
    def __init__(
        self,
        src: Union[str, Path, List[str]],
        doc_path: Union[str, Path, List[str]],
        model: Model,
        config: Optional[AppConfig] = None,
    ) -> None:
        """
        Constructor

        Args:
            src (Union[str, Path, List[str]]): Source File, this is usually is a file from build dir.
            doc_path (Union[str, Path, List[str]]): Path to resource LibreOffice document.
            model (ModelExample): Model for example.
            config (Optional[AppConfig], optional): App Config. Defaults to None.
        """
        self._src = util.get_path(src, ensure_absolute=True)
        self._doc_path = util.get_path(doc_path, ensure_absolute=True)
        self._config: AppConfig = config
        self._model = model
        if self._config is None:
            self._config = util.get_app_cfg()

    def embed(self) -> None:
        """
        Embeds python script in a LibreOffice Document and places doc in build dir.
        """
        # LO documents are zip file so this method unzip, adds script, updated manifest, rezips
        # and copies into build dir.
        def unzip(source: Path, dest: Path) -> None:
            with zipfile.ZipFile(str(source), "r") as zip_ref:
                zip_ref.extractall(str(dest))

        def zipdir(unzipped_path: Path, ziph: zipfile.ZipFile):
            for root, dirs, files in os.walk(unzipped_path):
                if len(files) == 0:
                    # it is a bit tricky adding empty dir to zip.
                    # this block add empty dirs using ZipInfo
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

        def get_odt_dest_path(tmp_dir: Path) -> Path:
            # path in tmpdir to zib version of LO document.
            src = self._doc_path
            dst = Path(tmp_dir, f"{src.stem}.zip")
            return dst

        def copy_script_to_unzipped(script_src: Path, zip_extract_dst: Path) -> None:
            p_script = zip_extract_dst / "Scripts" / "python"
            dst = Path(p_script, script_src.name)
            util.mkdirp(p_script)
            shutil.copy2(script_src, dst)

        def copy_zipped_to_build(zip_file: Path) -> None:
            build_path = build_util.get_build_path()
            util.mkdirp(build_path)
            build_dest = Path(build_path, f"{self._model.args.output_name}{self._doc_path.suffix}")
            shutil.copy2(zip_file, build_dest)

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

            zip_path = cp.dst_path
            zip_extract_dst = Path(p_tmp, zip_path.stem)
            unzip(source=zip_path, dest=zip_extract_dst)

            # update manifest in unziped dir
            manifest_path = zip_extract_dst / 'META-INF' / 'manifest.xml'
            mfs = ManifestScript(manifest_path=manifest_path, script_name=self._src.name, config=self._config)
            mfs.write()
            
            # remove zip file in tmp dir.
            if zip_path.exists():
                os.remove(zip_path)

            copy_script_to_unzipped(script_src=self._src, zip_extract_dst=zip_extract_dst)
            
            # zip unzipped dir with the new embeded script files
            zip_dest = cp.dst_path.parent / f"{self._model.args.output_name}.zip"
            # zip_dir(unzipped_path=zip_extract_dst, dst_zip=cp.dst_path)
            zip_dir(unzipped_path=zip_extract_dst, dst_zip=zip_dest)
            
            copy_zipped_to_build(zip_dest)
            # tmp dir will not del.

    def _validate(self) -> None:
        if self._src.exists() is False:
            raise FileNotFoundError(str(self._src))
