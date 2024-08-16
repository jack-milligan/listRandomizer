import random
import pandas as pd
import openpyxl


def read_excel_column(file_path: str, column_name: str) -> list:
    """
    Reads a specified column from an Excel file and returns a list of its values.

    Args:
    - file_path: str : Path to the Excel file.
    - column_name: str : Name of the column to be extracted.

    Returns:
    - list : List of values from the specified Excel column.
    """
    df = pd.read_excel(file_path)
    input_list = df[column_name].dropna().tolist()  # Drop NaN values and convert to list
    return input_list


def shuffle_list(input_list: list) -> list:
    """
    Shuffles a given list randomly.

    Args:
    - input_list: list : The list to shuffle.

    Returns:
    - list : Shuffled list.
    """
    random.shuffle(input_list)
    return input_list


def rank_list(input_list: list) -> list:
    """
    Assigns ranks to a list of values.

    Args:
    - input_list: list : The list of values to rank.

    Returns:
    - list : List of tuples, where each tuple contains a rank and the corresponding value.
    """
    return [(i + 1, value) for i, value in enumerate(input_list)]


def display_ranked_output(ranked_list: list) -> None:
    """
    Displays the ranked output in a readable format.

    Args:
    - ranked_list: list : List of ranked tuples (rank, value).
    """
    for rank, value in ranked_list:
        print(f"Rank {rank}: {value}")


# Main function to orchestrate the reading, shuffling, ranking, and displaying
def process_excel_column(file_path: str, column_name: str) -> None:
    """
    Processes the Excel column by reading, shuffling, ranking, and displaying the values.

    Args:
    - file_path: str : Path to the Excel file.
    - column_name: str : Name of the column to be processed.
    """
    # Read the Excel column
    input_list = read_excel_column(file_path, column_name)

    # Shuffle the list
    shuffled_list = shuffle_list(input_list)

    # Rank the list
    ranked_list = rank_list(shuffled_list)

    # Display the ranked output
    display_ranked_output(ranked_list)


# Main Method:
if __name__ == "__main__":
    file_path = '2024 Fantasy Owners.xlsx'  # Specify the path to your Excel file or upload local file to project
    column_name = 'Team Owners'  # Specify the column name

    # Process the Excel column (read, shuffle, rank, and display)
    process_excel_column(file_path, column_name)
