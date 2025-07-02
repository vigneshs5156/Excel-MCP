# 🤖 Claude Excel Assistant (via FastMCP)

Connect your Excel files to Claude (Anthropic LLM) using custom MCP tools — and control your spreadsheets with simple natural language.  
Forget formulas. Just ask!

---

## 🔍 What is this?

This project bridges **Claude AI** with **Excel/CSV files** through a powerful `FastMCP` server, allowing you to run common Excel operations like:

- Getting a quick overview of your data
- Filtering by column value
- Viewing column names
- Generating pivot tables
- Performing aggregations (sum, mean, etc.)
- Reading first N rows
- Listing available data files in a folder

All triggered by **natural language**, thanks to Claude's language model.

---

## 📂 Features

| Feature                    | Description                                                       |
|---------------------------|-------------------------------------------------------------------|
| ✅ Natural Language Input | Use Claude to describe your task in plain English                |
| 📈 Aggregation Support    | Sum, mean, min, max and more                                      |
| 📊 Pivot Tables           | Create pivot tables dynamically                                   |
| 📋 Column Extraction      | List all column names from your dataset                          |
| 🔎 Data Preview           | Read top N rows for quick inspection                             |
| 🔁 Filtering              | Filter rows by column value                                       |
| 📑 Overview               | Get statistical summary using `describe()`                       |
| 📂 File Browser           | View all data files in a folder                                   |

---

## ⚙️ Tools & Tech Stack

- [FastMCP](https://github.com/modelcontext/mcp) — lightweight framework to expose tools to LLMs
- `pandas` — Excel/CSV reading, filtering, aggregation, etc.
- `Claude` (via your preferred interface or API) — language model for natural query interpretation
- Python 3.10+

---

## 🛠️ MCP Tools Included

### `read_excel(file_path, num_rows)`
Reads the first `n` rows from a file.

---

### `get_column_names(file_path)`
Returns the list of column names in the dataset.

---

### `get_file_names(folder_path)`
Lists all files in a given directory (CSV/Excel supported).

---

### `filter_df(file_path, column_name, value)`
Filters a DataFrame by a specific column and value.

---

### `pivot_table(file_path, index, values, aggfunc)`
Creates a pivot table based on index and aggregation type.

---

### `get_agg(file_path, column_name, aggfunc)`
Performs sum, mean, max, etc., on a given column.

---

### `get_overview(file_path)`
Returns statistical summary (`describe()` output).

---

### `get_info(file_path)`
Returns technical info like data types and null counts (via `df.info()`).

---

## 🧪 Example Natural Language Prompts for Claude

- _“Show me the average sales by region.”_  
- _“List all products where quantity is greater than 100.”_  
- _“Create a pivot by category and sum the revenue.”_  
- _“Get the top 10 rows from the file ‘orders.csv’.”_  
- _“What are the column names in ‘sales.xlsx’?”_

---

## ✅ Benefits

- 📉 No more memorizing Excel functions
- 💬 Use natural conversation instead of formulas
- ⚡ Faster data analysis, even for non-tech users
- 🧠 Claude understands context and intent

---

## 🚀 How to Use

1. Clone this repository
2. Install dependencies:
    ```bash
    pip install pandas mcp
    ```
3. Start your MCP server:
    ```bash
    python your_script_name.py
    ```
4. Connect Claude to the MCP endpoint (e.g., via Claude's tool interface or a wrapper).

---

## 📌 Future Plans

- Claude-integrated Streamlit interface (work in progress)
- Error handling and dynamic column suggestions
- Output formatting improvements
- File upload support via web UI

---

## 👨‍💻 Author

Built with ❤️ by Vignesh S

Feel free to fork, improve, or integrate into your own LLM toolkits.

