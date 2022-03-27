# coding: utf-8
from pydantic import BaseModel, Field
from typing import Optional, List


class ExampleArgs(BaseModel):
    src_file: str
    embed_in_doc: bool
    document_name: Optional[str] = None
    include_modules: List[str] = Field(default_factory=list)
    include_paths: List[str] = Field(default_factory=list)
    remove_modules: List[str] = Field(default_factory=list)