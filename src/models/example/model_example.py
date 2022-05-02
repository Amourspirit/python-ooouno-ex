# coding: utf-8
from pydantic import BaseModel, validator
from typing import List, Optional
from .args import ExampleArgs
from .. import validators
from ..model_type_enum import ModelTypeEnum
from ...lib.enums import AppTypeEnum

class ModelExample(BaseModel):
    id: str
    version: str
    type: ModelTypeEnum
    app: AppTypeEnum
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
    