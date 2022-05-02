# coding: utf-8
from pydantic import BaseModel, Field, validator
from typing import List
from .. import validators

class ExampleArgs(BaseModel):
    src_file: str
    output_name: str
    include_modules: List[str] = Field(default_factory=list)
    include_paths: List[str] = Field(default_factory=list)
    remove_modules: List[str] = Field(default_factory=list)
    single_script: bool = False
    _str_null_empty = validator('src_file', 'output_name', allow_reuse=True)(
        validators.str_null_empty)