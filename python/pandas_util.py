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

DataFrame Aggregator Funcs
--------------------------
Functions to .groupby and .agg multiple subsets of a single DataFrame. Subsets are
defined by filtering certain columns within the df. All valid combinations of subsets
are created and given the same .groupby.agg process.

NamedTotal:\n
\tDataclass that defines necessary fields for filtering and renaming sub/grand totals
define_aggregator_from_dict:\n
\tDefines a function that reduces a DataFrame into a Series of measures.
dataframe_combination_agg:\n
\tCombine results of the same groupby - agg process on subsets of a df
"""

from dataclasses import dataclass
from enum import Enum
from itertools import combinations
from typing import List, Optional, Union, Callable, Any, Tuple

from pandas import concat, DataFrame, Series


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

# *************************************************************************************
# region COMBINATION AGGREGATOR


@dataclass
class NamedTotal:
    """Defines necessary fields for filtering and renaming sub/grand totals"""

    column: str
    alias: str
    selector: Callable[[Any], bool]


SubsetType = Tuple[NamedTotal]


def define_aggregator_from_dict(measures: dict):
    """Defines a function that reduces a DataFrame into a Series.

    The aggregator function uses a mapping of {str: Any} where str becomes the
    column name and Any is the aggregated value. Users must ensure their mapping
    reduces the input DataFrame as intended.

    Args
    ----
    measures: Mapping of column name to any value. Values are typically a single scalar
    reduction of a DataFrame column, but can be any obj.

    Returns
    -------
    Callable[[DataFrame], Series]
    """

    def aggregator(grp: DataFrame) -> Series:
        d = {}
        for alias, func in measures.items():
            d[alias] = func(grp)
        return Series(d, index=d.keys())

    return aggregator


def _generate_valid_subsets(
    totals: List[NamedTotal], required: List[str] = None
) -> SubsetType:
    """Yield a tuple of all NamedTotals that comprise a valid combination.

    Args
    ----
    totals: list of NamedTotal instances that define the subsets.
    required: list of column names that will only output aggregations that include
    total/subtotal selections.

    Yields
    -------
    Tuple of NamedTotal instances for a single aggregation.
    """

    def validator(subset) -> bool:
        """Given subset contains no duplicated cols and all required cols"""
        columns_in_subset = [c for c, _, _ in subset]
        is_distinct = len(columns_in_subset) == len(set(columns_in_subset))
        has_required = True
        if required is not None:
            has_required = set(required).issubset(columns_in_subset)

        return is_distinct & has_required

    subsets = []
    for combin in [combinations(totals, r) for r in range(1, len(totals) + 1)]:
        for subset in combin:
            subsets.append(subset)
    for valid in filter(validator, subsets):
        yield valid


def _create_subset_frame(df: DataFrame, subset: SubsetType) -> DataFrame:
    """Apply the subset definition to filter and transform the DataFrame

    Args
    ----
    df: DataFrame to be filtered and transformed
    subset: Definition of the column filters and alias values

    Returns
    -------
    Selected rows from df as defined by the selector functions in subset

    """
    for col, alias, selector in subset:
        df = df.loc[df[col].map(selector)].assign(**{col: alias})
    return df


def _combine_to_csv(filepath: str):
    """Set .csv filepath for outputting all dataframe aggregations

    The first agg output will add the header and replace any existing file. Subsequent
    aggs only append data to the file.
    """
    csv_kwargs = {"path_or_buf": filepath, "header": True, "mode": "w"}

    def append_df(df: DataFrame) -> None:
        df.to_csv(**csv_kwargs)
        if csv_kwargs["header"]:
            csv_kwargs["header"] = False
            csv_kwargs["mode"] = "a"

    return append_df


def _combine_to_memory(aggs: list):
    """Append dataframe aggregations to the given list instance."""

    def append_df(df) -> None:
        aggs.append(df)

    return append_df


def dataframe_combination_agg(
    df: DataFrame,
    groupby: List[str],
    totals: List[NamedTotal],
    aggregator: Callable[[DataFrame], Series],
    totals_only: List[str],
    csv_output_path: str = None,
) -> DataFrame:
    """Combine results of the same groupby - agg process on subsets of a df

    Args
    ----
    df: DataFrame to be aggregated\n
    groupby: Column names to groupby\n
    totals: List of NamedTotals which defines the total/subtotals in the combination
    aggs\n
    aggregator: Callable that transforms a single grouped dataframe into a series of
    measures.\n
    totals_only: List of column names that will only include their alias values defined
    in totals in the combination output.\n
    csv_output_path: Optional path to a csv file to output each aggregation instead of
    storing in memory. Useful when there are many combinations being created on a
    large dataframe as only the df and one aggregation are in memory at any time.\n

    Returns
    -------
    Original dataframe when output to csv, otherwise the result of applying and
    combining all aggregations.
    """
    if csv_output_path is None:
        fn_combin = _combine_to_memory(aggregations := [])
        fn_output = lambda: concat(aggregations)
    else:
        fn_combin = _combine_to_csv(csv_output_path)
        fn_output = lambda: df

    for subset in _generate_valid_subsets(totals, totals_only):
        (
            _create_subset_frame(df, subset)
            .groupby(groupby)
            .apply(aggregator)
            .pipe(fn_combin)
        )

    return fn_output()


# endregion
# *************************************************************************************
