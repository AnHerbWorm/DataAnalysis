""" 
TODO
    - [ ]   properties to show columns by grand total, sub total, group by, measures
            where would this lookup be useful?
            how to format the property output?
    - [ ]   see if worth it to refactor totals into {
            'grand': {current self.grandtotals},
            'sub': {current self.subtotals}
            } and make grand/sub properties that extract from this private dict
    - [ ]   static method to get column names from {alias: NamedTuple} format
            eg _col_specs_to_list_(dict, attr_name, distinct=True):
                lst = [specs.get(attr_name) for specs in dict.values()]
                if distinct:
                    return list(dict.fromitems(lst))
                else:
                    return lst
            useful for cleaning up some of the longer list comprehensions?
"""
# standard modules
import logging
from collections import namedtuple, defaultdict
from itertools import combinations

# third-party modules
import pandas as pd

_TotalColSpecs = namedtuple("TotalColSpecs", ["column", "sub_values", "incl"])
_MeasureColSpecs = namedtuple("MeasureColSpecs", ["column", "aggfunc"])


class DataFrameAggregator:
    """DataFrameAggregator class to calculate measures on a DataFrame with grand/sub total column aggregations.

    Any column in the frame can have:
        0 or 1 Grandtotals. Raises KeyError if multiple grandtotals are added for the same column\n
        0 or more Subtotals. Raises KeyError if multiple subtotals attempt to use the same alias\n
        0 or more Measure calculations. Raises KeyError if multiple measures attempt to use the same name\n
    Creating a grand/sub total requires that column in the group_cols and are auto added\n
    Creating a measure column does not require that column in the group_cols\n
    """

    def __init__(self, frame: pd.DataFrame = None, friendly_name: str = "Unnamed"):
        """Create a DataFrameAggregator obj that acts on the frame to produce measures

        The obj does nothing when instantiated and requires groupby/grandtotal/subtotal/measures to be added\n

        Args
        ----
        frame: Optional. The dataframe to aggregate, if omitted the DataFrameAggregator validates itself
        once a DataFrame is added via set_frame()
        friendly_name: Optional. Gives this aggregator a name in __repr__ output
        """
        self.friendly_name = friendly_name
        self.frame = frame
        if self.frame is not None:
            self._frame_cols = list(frame.columns)
        else:
            self._frame_cols = []
        # list of columns to groupby during aggregate
        self.group_cols = []
        # dict of {alias: TotalColSpecs namedtuple}
        self.grandtotals = {}
        self.subtotals = {}
        # dict of {alias: MeasureColSpecs namedtuple}
        self.measures = {}

    def __repr__(self):
        return (
            f"DataFrameAggregator ({self.friendly_name}) with: "
            f"{len(self.grandtotals)} grand totals, "
            f"{len(self.subtotals)} sub totals, "
            f"{len(self.measures)} measures"
        )

    def set_frame(self, frame: pd.DataFrame) -> None:
        """Set the frame that this aggregator will act upon

        Args
        ----
        frame: The dataframe to aggregate. Adds a reference to the original frame so any changes made
        to it will be reflected in this instance

        Raise
        -----
        Errors raise if any of the total or measure columns already set do not exist in
        the dataframe being set
        """
        preset_columns = list(
            set(self.group_cols + [specs.column for specs in self.measures.values()])
        )
        frame_columns = frame.columns
        for col in preset_columns:
            if col not in frame_columns:
                raise ValueError(
                    f"{col} was set to be aggregated but is not part of the given dataframe!"
                )
        self.frame = frame

    def add_groupby_column(self, column: str) -> None:
        """Include column in the groupby selection when aggregating, without making a grand/sub total

        Args
        ----
        column: name of the column to groupby

        Raise
        -----
        KeyError if the column does not exist in the DataFrame
        """
        if self._column_in_frame_(column):
            self.group_cols.append(column)
            # ensure the list is only unique columns, in the same order they were entered
            self.group_cols = list(dict.fromkeys(self.group_cols))

    def add_grandtotal(self, column: str, alias: str) -> None:
        """Include column in the groupby selection when aggregating, and make a grand total

        Args
        ----
        column: name of the column to groupby
        alias: value to describe the grand total for this column

        Raise
        -----
        KeyError if the column does not exist in the DataFrame
        KeyError if the column is assigned more than one grandtotal
        KeyError if the alias is used more than once (for any column, total or measure)
        """
        self.add_groupby_column(column)
        if column in self.grandtotals.values():
            raise KeyError(f"{column} already has a grand total assigned!")
        if self._alias_is_unique_(alias):
            self.grandtotals[alias] = _TotalColSpecs(column, False, True)

    def add_subtotal(
        self, column: str, alias: str, subtotal_filter: list, subtotal_incl: bool = True
    ) -> None:
        """Include column in the groupby selection when aggregating, and make a subtotal

        Values used in the subtotal are defined by subtotal_filter. Either by explicity inclusion (default)
        or excluded (where every other value in the column forms the subtotal)

        Args
        ----
        column: name of the column to groupby\n
        alias: value to describe the sub total for this column\n
        subtotal_filter: list of values to form the subtotal, how is set by subtotal_incl\n
        subtotal_incl: Default True. Sets the behaviour for filtering the DataFrame by the values in subtotal_filter.
        True selects only the values listed, False excludes only the values listed.

        Raise
        -----
        KeyError if the column does not exist in the DataFrame
        KeyError if the alias is used more than once (for any column, total or measure)
        """
        self.add_groupby_column(column)
        if self._alias_is_unique_(alias):
            self.subtotals[alias] = _TotalColSpecs(
                column, subtotal_filter, subtotal_incl
            )

    def add_measure(self, column: str, alias: str, aggfunc) -> None:
        """Add a measure to the aggregation

        Args
        ----
        column: name of the column to aggregate into a measure
        alias: header to use for the measure column
        aggfunc: function to aggregate with. Accepts any value that works in DataFrameGroupBy.agg()

        Raise
        -----
        KeyError if the column does not exist in the DataFrame
        KeyError if the alias is used more than once (for any column, total or measure)
        """
        if self._column_in_frame_(column) and self._alias_is_unique_(alias):
            self.measures[alias] = _MeasureColSpecs(column, aggfunc)

    def aggregate(self, totals_only: list = None) -> pd.DataFrame:
        """Perform all aggregations and return them in a single concatenated DataFrame

        Args
        ----
        totals_only: Optional. List of column names to only output grand/sub totals on.
        Other groupby columns which are excluded will use all distinct values as groups,
        as per the typical behaviour.

        Return
        ------
        DataFrame of all aggregations concatenated
        """
        logging.info(f"START-- .aggregate() for {self}")
        self._can_perform_aggregate_()

        all_stats = []

        all_totals = {**self.grandtotals, **self.subtotals}
        # maximum items in combinations = count of distinct columns with either grand or sub total applied
        comb_max = len(set([specs.column for specs in all_totals.values()]))
        if totals_only:
            # check that the columns listed in totals_only are viable
            cols_in_totals = list(set([specs.column for specs in all_totals.values()]))
            for col in totals_only:
                if col not in cols_in_totals:
                    raise ValueError(
                        f"{col} in 'totals_only' does not have a grand/sub total set. Expected one of {cols_in_totals}"
                    )

            # calculate only total aggregations on certain columns
            # minimum items in combination = 1 + count of totals_only
            comb_min = min(comb_max, 1 + len(totals_only))
        else:
            # calculate all aggregations
            # minimum items in combinations = 1
            # compute measures on frame without any grand/sub total transformations applied
            # TODO - see if this will work with minimum zero instead; if it needs a special rule in the subset this method is preferred
            comb_min = 1
            all_stats.append(
                self._compute_measures_(self.frame, self.group_cols, self.measures)
            )

        for r in range(comb_min, comb_max + 1):
            for subset in combinations(all_totals.keys(), r):
                # subset is a list of aliases to add to the measures computations
                # subset_columns rebuilds the dict of {alias: _TotalColSpecs} for only aliases in this subset
                # parent_columns is a list of all the columns that would be affected by this subset (duplicates allowed)
                subset_columns = {
                    alias: specs
                    for alias, specs in all_totals.items()
                    if alias in subset
                }
                parent_columns = [specs.column for specs in subset_columns.values()]

                # the combinations do not account for multiple totals using the same parent column. These subsets are
                # ignored as it is not possible to do multiple totals for 1 column in the same aggregation
                # any subsets that do not contain an alias for all columns in totals_only are also ignored
                parent_columns_distinct = len(parent_columns) == len(
                    set(parent_columns)
                )
                has_all_total_cols = True
                if totals_only:
                    has_all_total_cols = set(totals_only).issubset(parent_columns)

                # process the subset aggregations
                if parent_columns_distinct and has_all_total_cols:
                    # grandtotals only change the values in the parent column to aliases
                    grandtotals = {
                        specs.column: alias
                        for alias, specs in subset_columns.items()
                        if specs.sub_values is False
                    }
                    df = self.frame.transform(
                        self._apply_alias_transform_, "index", to_transform=grandtotals
                    )

                    # subtotals first filter the column then transform all values to the aliases
                    subtotals = {
                        alias: specs
                        for alias, specs in subset_columns.items()
                        if isinstance(specs.sub_values, list)
                    }
                    for specs in subtotals.values():
                        df = df[df[specs.column].isin(specs.sub_values) == specs.incl]
                    df = df.transform(
                        self._apply_alias_transform_,
                        "index",
                        to_transform={
                            specs.column: alias for alias, specs in subtotals.items()
                        },
                    )

                    # measures can now be calculated on the dataframe as it has been adjusted for grand/sub totals
                    all_stats.append(
                        self._compute_measures_(df, self.group_cols, self.measures)
                    )

        logging.info(f"END---- .aggregate() for {self}")
        return pd.concat(all_stats)

    @staticmethod
    def _apply_alias_transform_(col, to_transform: dict) -> pd.DataFrame:
        """Transform all records in to_transform.keys() columns to the value for that key"""
        return to_transform.get(col.name, col)

    @staticmethod
    def _compute_measures_(
        frame: pd.DataFrame, group_cols: list, measures: dict
    ) -> pd.DataFrame:
        """Calculate measures on the given frame using measures specified in the class attr

        Args
        ----
        frame: Optional (default self.frame). DataFrame to aggregate
        group_cols: list of column names to groupby
        measures: dict of {'alias': MeasureSpecs(column, aggfunc)}. Designed to be used with measures attr of a
        DataFrameAggregator instance, which is a properly structured dict for this method

        Return
        ------
        New DataFrame after aggregations are done
        """
        agg_dict = defaultdict(list)
        for specs in measures.values():
            agg_dict[specs.column].append(specs.aggfunc)
        agg = frame.groupby(group_cols).agg(agg_dict)
        # rename measure columns
        agg.columns = list(measures.keys())

        return agg

    def _column_in_frame_(self, column: str) -> bool:
        """raise error if the column is not in this obj's frame"""
        if self.frame is None:
            return True
        elif column in self._frame_cols:
            return True
        else:
            raise ValueError(f"'{column}' is not a valid column within {self.frame}")

    def _alias_is_unique_(self, alias: str) -> bool:
        """raise a KeyError if the alias is already in use.

        Duplicate aliases are not supported by the aggregation logic
        """
        if any(
            [alias in self.grandtotals, alias in self.subtotals, alias in self.measures]
        ):
            raise KeyError(
                f"{alias} already in use! Duplicate aliases are not supported."
            )
        return True

    def _can_perform_aggregate_(self):
        """raise an error if the state of attrs does not allow for aggregation"""
        if self.frame is None:
            raise ValueError(f"No dataframe has been set! Use .set_frame() to add one.")
        if len(self.group_cols) == 0:
            raise ValueError(f"{self} has no groupby columns")
        if len(self.measures) == 0:
            raise ValueError(f"{self} has no aggregation measures to calculate")
