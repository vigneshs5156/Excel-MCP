import pandas as pd
import os
from mcp.server.fastmcp import FastMCP

mcp = FastMCP("Excel-MCP")

def convert_valid_column_name(column_name):
    """
    Converts a given column name into a valid format by:
    - Replacing spaces with underscores
    - Converting all characters to lowercase

    Parameters:
        column_name (str): The original column name.

    Returns:
        str: A cleaned, standardized column name.

    Raises:
        ValueError: If the input is not a string.
    """
    if not isinstance(column_name, str):
        raise ValueError("Column name must be a string.")

    column_name = "_".join(column_name.split())  # Replace spaces with underscores
    column_name = column_name.lower()            # Convert to lowercase
    return column_name


@mcp.tool()
def read_excel(file_path, num_rows=5):
    """
    Reads the first few rows of a file (Excel or CSV).

    Parameters:
        file_path (str): Path to the input file.
        num_rows (int): Number of rows to display. Default is 5.

    Returns:
        pd.DataFrame: A DataFrame containing the first `num_rows` rows.
    """
    df = pd.read_excel(file_path) if file_path.endswith('.xlsx') else pd.read_csv(file_path)
    return df.head(num_rows)


@mcp.tool()
def get_column_names(file_path):
    """
    Retrieves the column names from the provided Excel or CSV file.

    Parameters:
        file_path (str): Path to the input file.

    Returns:
        list: A list of column names.
    """
    df = pd.read_excel(file_path) if file_path.endswith('.xlsx') else pd.read_csv(file_path)
    return list(df.columns)


@mcp.tool()
def get_file_names(folder_path):
    """
    Lists all file names in a given folder (ignores directories).

    Parameters:
        folder_path (str): Path to the directory.

    Returns:
        list: A list of filenames found in the folder.
    """
    return [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]


@mcp.tool()
def filter_df(file_path, column_name, value):
    """
    Filters a DataFrame based on a column's value.

    Parameters:
        file_path (str): Path to the Excel or CSV file.
        column_name (str): Column to filter on (must match exactly).
        value (str|int|float): The value to match in the given column.

    Returns:
        pd.DataFrame: A filtered DataFrame containing only matching rows.

    Note:
        The column name must be provided exactly as it appears in the file.
    """
    df = pd.read_excel(file_path) if file_path.endswith('.xlsx') else pd.read_csv(file_path)
    df = df.query(f"{column_name} == {value}") 
    return df


@mcp.tool()
def pivot_table(file_path, index, values, aggfunc='mean'):
    """
    Creates a pivot table from the Excel or CSV data.

    Parameters:
        file_path (str): Path to the Excel or CSV file.
        index (str): Column name to group by (rows of pivot table).
        values (str): Column name whose values are aggregated.
        aggfunc (str): Aggregation function (e.g., 'sum', 'mean', 'max'). Default is 'mean'.

    Returns:
        pd.DataFrame: A pivot table summarizing the data.
    """
    df = pd.read_excel(file_path) if file_path.endswith('.xlsx') else pd.read_csv(file_path)
    pivot_table = df.pivot_table(index=index, values=values, aggfunc=aggfunc)
    return pivot_table


@mcp.tool()
def get_agg(file_path, column_name, aggfunc='sum'):
    """
    Performs an aggregation operation on a specific column.

    Parameters:
        file_path (str): Path to the Excel or CSV file.
        column_name (str): The name of the column to aggregate.
        aggfunc (str): Aggregation function (e.g., 'sum', 'mean', 'min', 'max').

    Returns:
        float|int: The aggregated result.

    Raises:
        ValueError: If the column name does not exist or is not numerical.
    """
    column_name = convert_valid_column_name(column_name)
    df = pd.read_excel(file_path) if file_path.endswith('.xlsx') else pd.read_csv(file_path)

    # Check if the converted column name exists in DataFrame
    if column_name not in map(convert_valid_column_name, df.columns):
        raise ValueError(f"Column '{column_name}' not found in file.")

    return df[column_name].agg(aggfunc)

@mcp.tool()
def get_overview(file_path):
    df = pd.read_excel(file_path) if file_path.endswith('.xlsx') else pd.read_csv(file_path)

    return df.describe()

@mcp.tool()
def get_info(file_path):
    df = pd.read_excel(file_path) if file_path.endswith('.xlsx') else pd.read_csv(file_path)
    return df.info()

