# Licensed to the Apache Software Foundation (ASF) under one
# or more contributor license agreements.  See the NOTICE file
# distributed with this work for additional information
# regarding copyright ownership.  The ASF licenses this file
# to you under the Apache License, Version 2.0 (the
# "License"); you may not use this file except in compliance
# with the License.  You may obtain a copy of the License at
#
#   http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing,
# software distributed under the License is distributed on an
# "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
# KIND, either express or implied.  See the License for the
# specific language governing permissions and limitations
# under the License.
import io
from typing import Any

import pandas as pd

from superset.utils.core import GenericDataType


def quote_formulas(df: pd.DataFrame) -> pd.DataFrame:
    """
    Make sure to quote any formulas for security reasons.
    """
    formula_prefixes = {"=", "+", "-", "@"}

    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].apply(
            lambda x: (
                f"'{x}"
                if isinstance(x, str) and len(x) and x[0] in formula_prefixes
                else x
            )
        )

    return df


def df_to_excel(df: pd.DataFrame, **kwargs: Any) -> Any:
    output = io.BytesIO()

    # make sure formulas are quoted, to prevent malicious injections
    df = quote_formulas(df)

    # pylint: disable=abstract-class-instantiated
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, **kwargs)

        workbook = writer.book
        worksheet = writer.sheets[kwargs.get("sheet_name", "Sheet1")]
        
        # Create number format with thousand separators
        number_format = workbook.add_format({"num_format": "#,##0.00"})
        
        has_index = kwargs.get("index", True)
        index_offset = 1 if has_index else 0
        
        # Format numeric columns and set column widths
        for idx, col in enumerate(df.columns):
            col_idx = idx + index_offset
            is_numeric = pd.api.types.is_numeric_dtype(df[col])
            
            # Write numeric values with formatting
            if is_numeric:
                for row_idx, value in enumerate(df[col], start=1):
                    if pd.notna(value):
                        try:
                            worksheet.write_number(row_idx, col_idx, float(value), number_format)
                        except (ValueError, TypeError):
                            pass  # Keep the original value from df.to_excel
            
            # Auto-size columns
            max_len = len(str(col))
            for value in df[col].astype(str):
                max_len = max(max_len, len(value))
            
            worksheet.set_column(col_idx, col_idx, min(max_len+8, 50))

    return output.getvalue()


def apply_column_types(
    df: pd.DataFrame, column_types: list[GenericDataType]
) -> pd.DataFrame:
    """
    Applies the column types to the dataframe to prepare for an excel export

    :param df: The dataframe to apply the column types to
    :param column_types: The types of the columns
    :return: The dataframe with the column types applied
    """
    for column, column_type in zip(df.columns, column_types, strict=False):
        if column_type == GenericDataType.NUMERIC:
            try:
                df[column] = pd.to_numeric(df[column])
                # if the number is too large, convert it to a string
                # Excel does not support numbers larger than 10^15
                df[column] = df[column].apply(
                    lambda x: str(x)
                    if isinstance(x, (int, float)) and abs(x) > 10**15
                    else x
                )
            except ValueError:
                df[column] = df[column].astype(str)
        elif pd.api.types.is_datetime64tz_dtype(df[column]):
            # timezones are not supported
            df[column] = df[column].astype(str)
    return df
