from abc import ABC, abstractmethod
from os import PathLike
from typing import Any


class TemplateReader[T](ABC):
    @abstractmethod
    def read(self, file: str | PathLike[str]) -> T: ...


class TemplateWriter(ABC):
    @abstractmethod
    def write(self, vars: dict[str, Any], file: str | PathLike[str]) -> None: ...
