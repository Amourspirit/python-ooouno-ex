# coding: utf-8
import verr
from pydantic import BaseModel, validator
from typing import List, Optional
from .args import ExampleArgs
from .. import validators

class ModelExample(BaseModel):
    id: str
    version: str
    name: str
    args: ExampleArgs
    methods: List[str]
    _str_null_empty = validator('name', allow_reuse=True)(
        validators.str_null_empty)
    
    @validator('id')
    def validate_id(cls, value: str) -> str:
        if value == 'ooouno-ex':
            return value
        raise ValueError(
            f"root/id/ must be 'ooouno-ex'. Current value: {value}")
    
    @validator('version')
    def validate_version(cls, value: str) -> str:
        v_result = verr.Version.try_parse(value)
        if v_result[0] == False:
            raise ValueError(f"root/version has a bad format: {value}")
        else:
            cls.ver = v_result[1]
        return value