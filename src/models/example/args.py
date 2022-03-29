# coding: utf-8
from pydantic import BaseModel, Field, validator
from typing import Optional, List
from .. import validators

class ExampleArgs(BaseModel):
    src_file: str
    output_name: str
    embed_in_doc: bool
    document_name: Optional[str] = None
    include_modules: List[str] = Field(default_factory=list)
    include_paths: List[str] = Field(default_factory=list)
    remove_modules: List[str] = Field(default_factory=list)
    _str_null_empty = validator('src_file', 'output_name', allow_reuse=True)(
        validators.str_null_empty)