# coding: utf-8

import json
from pathlib import Path
import stickytape
import os
from dataclasses import dataclass
from typing import List, Optional
from distutils.sysconfig import get_python_lib
import fileinput
import re
from config import AppConfig
from .embed_py_script import EmbedScriptPy
from .copy_resource import CopyResource
from ..utils import util
from ..models.example.model_example import ModelExample

@dataclass
class BuilderArgs:
    config_json: str
    """
    Path to json file that contains configuratin
    """
    allow_print: bool = False
    """
    If True some print statements will be made in the terminal
    """
    config_app: Optional[AppConfig] = None
    """
    Optional AppConfig
    """
    embed_in_doc: bool = False
    """
    If True then an odt file with embeded script is generated
    """
    embed_doc: Optional[str] = None
    """
    The document to embed compile script into.
    
    Requires ``embed_in_doc`` be ``True``.
    
    If ``embed_in_doc`` be ``True`` then the default used document to embed
    is internal .odt file.
    """

class Builder:
    """Builder Class"""

    # region Constructor
    def __init__(self, args: BuilderArgs):
        """
        Constructor

        Arguments:
            args (BuilderArgs): Builder Args
        """
        self._config: AppConfig = args.config_app
        if self._config is None:
            self._config = util.get_app_cfg()
        self._allow_print = args.allow_print
        self._embed = args.embed_in_doc
        self._embed_doc = args.embed_doc
        self._dest_file = ""
        self._json_cfg = util.get_path(
            path=args.config_json.replace('\\', '/').split('/'),
            ensure_absolute=True)
        # self._json_cfg = util.get_path_from_lst(
        #     "src/examples/message_box/config.json".split("/"), ensure_absolute=True
        # )
        self._src_path = self._json_cfg.parent
        self._site_pkg_dir = None
        with open(self._json_cfg, "r") as file:
            jdata: dict = json.load(file)
        self._model = ModelExample(**jdata)
        self._src_file = self._get_src_file()

    # endregion Constructor

    def _get_src_file(self) -> Path:
        parts = self._model.args.src_file.replace("\\", "/").split("/")
        rel = util.get_path(parts, ensure_absolute=False)
        if rel.is_absolute():
            return rel
        return Path(self._src_path, rel)

    def _get_python_modules(self, add_flag="--add-python-module") -> str:
        if len(self._model.args.include_modules) == 0:
            return ""
        result = ""
        for val in self._model.args.include_modules:
            result = result + f" {add_flag} {val}"
        return result

    def _get_include_paths(self) -> List[str]:
        paths = [itm for itm in self._model.args.include_paths]
        paths.append(self.site_pkg_dir)
        return paths

    def _get_site_include_path(self) -> str:
        paths = self._get_include_paths()
        result = ""
        for i, p in enumerate(paths):
            if i > 0:
                result = result + " "
            result = f"{result}--add-python-path {p!r}"
        return result

    def _remove_modules(self) -> None:
        if len(self._model.args.remove_modules) == 0:
            return
        s_pattern_start = "^(?:[ \t]*)__stickytape_write_module\('"
        s_pattern_end = ".*$"
        # combine default config modules with example config remove modules
        r_mods = set(self._config.build_remove_modules + self._model.args.remove_modules)
        for i, p in enumerate(r_mods):
            pattern = f"{s_pattern_start}{p}{s_pattern_end}"
            matched = re.compile(pattern).search
            with fileinput.FileInput(
                self._dest_file, inplace=1, backup=f".{i}.bak"
            ) as file:
                for line in file:
                    if not matched(line):  # save lines that do not match
                        print(line, end="")  # this goes to filename due to inplace=1
    
    def _get_g_exported(self) -> str:
        # g_exportedScripts = (
        #     main,
        # )
        count = len(self._model.methods)
        result = "\ng_exportedScripts = ("
        if count > 0:
            body = ",".join(self._model.methods)
            result += body
        if count == 1:
            result += ","
        result += ")\n"
        return result
    
    def _append_g_exported(self) -> None:
        with open(self._dest_file, 'a') as file:
            file.write(self._get_g_exported())
    
    def _embed_script(self) -> None:
        if self._embed_doc is None:
            src_doc = self._config.app_res_blank_odt
        else:
            src_doc = util.get_path(self._embed_doc)
        cp = CopyResource(
            src=src_doc,
            dst=None,
            clear_prev=False,
            config=self._config,
            src_is_res = self._embed_doc is None
        )
        emb = EmbedScriptPy(
            src=self._dest_file,
            doc_path=cp.src_path,
            model=self._model,
            config=self._config
        )
        emb.embed()

    # region Public Methods
    def build(self) -> bool:
        """
        Builds the plugin into the dist folder outlined in settings.json
        @return: `True` of the build is a success; Otherwise, `False`
        """

        dist_dir = util.get_path(
            self._config.app_build_dir, ensure_absolute=True
        )
        dest_file = dist_dir / f"{self._model.args.output_name}{self._src_file.suffix}"
        self._dest_file = str(dest_file)

        if not os.path.exists(dist_dir):
            os.makedirs(dist_dir)
            if self.allow_print:
                print("Made Dir: " + dist_dir)

        if os.path.exists(self._dest_file):
            os.remove(self._dest_file)
            if self.allow_print:
                print("Deleted File: " + self._dest_file)
        # region Make file using stickytape
        # processer = "stickytape"
        # site_inc_path = self._get_site_include_path()
        # params = f"'{self._src_file}' {site_inc_path} --output-file {self._dest_file!r}".replace(
        #     "  ", " "
        # ).strip()
        # cmd = f"{processer} {params}"
        # if self.allow_print:
        #     print(cmd)

        # os.system(cmd)

        # sys.exit()
        output = stickytape.script(
            path=str(self._src_file),
            add_python_modules=[],
            add_python_paths=self._get_include_paths(),
        )
        with open(self._dest_file, "w") as output_file:
            output_file.write(output)
        # endregion Make file using stickytape
        # region Append Global Exports

        # endregion Append Global Exports
        # region Report
        if self.allow_print == True and os.path.exists(self._dest_file):
            print("Generated File: " + self._dest_file)
        # endregion Report
        self._remove_modules()
        self._append_g_exported()
        if self._embed:
            self._embed_script()
        return True

    # endregion Public Methods

    @property
    def site_pkg_dir(self):
        if self._site_pkg_dir is None:
            self._site_pkg_dir = get_python_lib()
        return self._site_pkg_dir

    @property
    def allow_print(self) -> bool:
        return self._allow_print

    @allow_print.setter
    def allow_print(self, value: bool) -> None:
        self._allow_print = value
