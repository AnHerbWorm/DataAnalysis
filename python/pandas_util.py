"""
pandas_util.py

Utility classes and functions for working with pandas library.

ColumnEnumerator
----------------
Class to provide attribute access to column names and common operations for
working with DataFrames in larger projects.
Smaller projects or .ipynb files will benefit more from the attribute access built-into
DataFrames.

Init from list of ColumnSpecs:
>>> Cols = ColumnEnumerator([ColumnSpecs(...)])
Get a single column (as Series):
>>> df[Cols.column_attr]
Select all columns in the ColumnEnumerator:
>>> df[Cols.select]
>>> [c1_name, c2_name, ...]
Get dtype mapping:
>>> Cols.dtype_mapping
>>> {'c1_name': 'c1_dtype', 'c2_name', 'c2_dtype', ...}
Group accessor (requires groups defined via Cols.add_group method):
>>> Cols.G.GROUP_NAME

"""

from dataclasses import dataclass
from enum import Enum
from typing import List, Optional, Union


class PandasDtype(str, Enum):
    """https://pandas.pydata.org/pandas-docs/stable/user_guide/basics.html#dtypes"""

    BOOLEAN = "boolean"
    CATEGORICAL = "category"
    DATETIME = "datetime64[ns]"
    DATETIMEZONE = "datetime64[ns, <tz>]"
    INT64 = "Int64"
    INT32 = "Int32"
    INT16 = "Int16"
    INT8 = "Int8"
    STRING = "string"
    UINT64 = "UInt64"
    UINT32 = "UInt32"
    UINT16 = "UInt16"
    UINT8 = "UInt8"


# *************************************************************************************
# region ColumnEnumerator


@dataclass
class ColumnSpecs:
    """Dataclass representation of a DataFrame columns

    Attributes
    ----------
    attr: Name of the attribute that will be used to access this column's name
    name: Name of the column
    dtype: String representation of a pandas dtype
    desc:  Optional field description, for info purposes only
    """

    attr: str
    name: str
    dtype: Union[PandasDtype, str]
    desc: Optional[str] = None

    def __post_init__(self):
        available_types = [e.value for e in PandasDtype]
        dtype_url = (
            "https://pandas.pydata.org/pandas-docs/stable/user_guide/basics.html#dtypes"
        )
        if isinstance(self.dtype, PandasDtype):
            self.dtype = self.dtype.value
        if self.dtype not in available_types:
            raise ValueError(
                f"'{self.dtype}' is not a pandas dtype string representation."
                f"See {dtype_url} for more details"
            )


class _GroupObject:
    """Do-nothing class that acts only as a holder for attributes"""


class _ColumnEnumerator:
    """Base class that does not add column lookup attributes.

    Intended for internal use to prevent infinitely recursive sub-grouping of columns
    in a ColumnEnumerator instance.
    """

    def __init__(self, columns: List[ColumnSpecs]):
        """

        Args
        ----
        columns: list of all ColumnSpecs instances to add to the enumerator
        """
        self._specs = columns

    @property
    def specs(self) -> List[ColumnSpecs]:
        """List of ColumnSpecs in this instance"""
        return self._specs

    @property
    def select(self) -> List[str]:
        """List of column names in the order given during init"""
        return [c.name for c in self._specs]

    @property
    def dtype_mapping(self) -> dict:
        """Dict of name:dtype pairs for all columns"""
        return {c.name: c.dtype for c in self._specs}

    def __repr__(self) -> str:
        specs = ",\n  ".join([repr(c) for c in self._specs])
        return f"_ColumnEnumerator(columns=[\n  {specs}])"


class ColumnEnumerator(_ColumnEnumerator):
    """Utility class for attribute access to column names and info."""

    def __init__(self, columns: List[ColumnSpecs]):
        super().__init__(columns)
        self._groups = None
        self._set_accessors()

    def __repr__(self) -> str:
        # only need to remove the leading _ from super class
        return super().__repr__()[1:]

    def _set_accessors(self):
        for col in self._specs:
            self.__setattr__(col.attr, col.name)

    @property
    def G(self):
        """Accessor for subgroups of columns within this group"""
        return self._groups

    def add_group(self, attr: str, columns: List[ColumnSpecs] | List[str]) -> None:
        """Add a group to the GroupableColumnEnumerator.g property

        Args
        ----
        attr: Attribute name for accessing the group (recommend UPPERCASE)
        columns: list of ColumnSpecs instances, or list of column names.
        """
        if self._groups is None:
            self._groups = _GroupObject

        columns = list(map(self._get_specs_instance, columns))
        self._groups.__setattr__(attr, _ColumnEnumerator(columns))

    def _get_specs_instance(self, value: Union[ColumnSpecs, str]) -> ColumnSpecs:
        """Return ColumnSpecs instance represented by the column name

        Column must exist within this instances specs attribute
        """
        if isinstance(value, str):
            value = [c for c in self._specs if c.name == value][0]

        if isinstance(value, ColumnSpecs):
            if value in self._specs:
                return value

        raise ValueError(
            f"'{value}' is not a name or ColumnSpecs instance within this"
            "ColumnEnumerator"
        )

    @classmethod
    def from_csv(cls, filepath: str, ignore_header: bool):
        """Create ColumnEnumerator instance with columns read from a .csv file.

        Args
        ----
        filepath: str filepath to a .csv file
        ignore_header: indicate if the file has a header row to ignore
        """
        with open(filepath, "r") as csv:
            rows = [row.strip().split(",") for row in csv.readlines()]
        if ignore_header:
            rows = rows[1:]
        cols = list(map(lambda row: ColumnSpecs(*row), rows))
        return cls(cols)


# endregion
# *************************************************************************************
