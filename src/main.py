import pandas as pd
import re
import csv
import numpy as np
import warnings
import xlsxwriter
import json
from tqdm import tqdm

from pandas.core import series
with open('data/dictionaries/default_hamlets.json') as f:
    default_hamlet = json.load(f)
with open('data/dictionaries/abbreviations.json') as f:
    abbreviation = json.load(f)
with open('data/dictionaries/latitude.json') as f:
    latitude = json.load(f)
with open('data/dictionaries/longitude.json') as f:
    longitude = json.load(f)

df_list = []
input_filepath = "data/input/Tomini.xlsx"
options = {
    'strings_to_formulas': False,
    'strings_to_urls': False,
    'strings_to_numbers': False
}

reference_rx = r"((\+|\=|\->|\?->)([0-9][0-9].([0-9][0-9][0-9])))"
""" 
Regular expression: Matches any sequence of characters that starts with '+', '=', '->', or '?->' followed by two digits, a period, and three more digits.
"""
reference2_rx = r"[0-9][0-9]\.[0-9][0-9][0-9]"
"""
Regular expression: Matches any sequence of characters that consists of two digits, a period, and three more digits.
"""
pages_rx = r"(p\d+\w|\d+\/(\d{1,}|\w)|p\d)"
"""
Regular expression: Pattern to match page numbers in text. Matches any sequence of characters that starts with 'p' followed by one or more digits, a letter, or a combination of digits and forward slashes.
"""
location1_rx = r"(\+?([A-Z]{3}|OU)(?<![A-Z]{4})(?![A-Z]):)"
"""
Regular expression: Matches any sequence of characters that starts with an optional '+' followed by three capital letters or the string 'OU', not immediately preceded by four capital letters, and ends with a colon.
"""
location2_rx = r"((?=\+).[A-Z]{3})"
"""
Regular expression: Matches any sequence of characters that starts with a '+' followed by three capital letters.
"""
translation_comment_rx = r'\(.*?\)|\".*?\"'
"""
Regular expression: Matches any sequence of characters enclosed in either parentheses or double quotes.
"""
editing_comment_rx = r"(\#(.*?)\#)"
"""
Regular expression: Matches any sequence of characters enclosed in hash symbols.
"""
grammar_rx = r"((?=\S*[-]|RDP|RDP\d|te=).*)"
"""
Regular expression: Matches any sequence of characters that contain a hyphen, 'RDP', 'RDP' followed by a digit, or 'te=' anywhere in the string.
"""
group_rx = r"^\[\S|\S\]$"
"""
Regular expression: Pattern to match groups. Matches any sequence of characters that starts and ends with square brackets and contains at least one non-whitespace character in between.
"""
gender_rx = r"(\(l-l\)|\(per\))"
"""
Regular expression: Pattern to match gender. Matches either the string '(l-l)' or '(per)'.
"""
register_rx = r"(\(K\)|\(H\))"
"""
Regular expression: Pattern to match register. Matches either the string '(K)' or '(H)'.
"""
remove_comma_rx = r"(\,\s+$)"
"""
Regular expression: Pattern for removing commas followed by one or more whitespaces at the end of a string.
"""
recorded_rx = r"\~"
"""
Regular expression: Pattern for matching the "~" symbol which indicates if a recording exists
"""
pronunciation_rx = r"\*w|\[w\]|\[ß\]|\[\^b\]|\+'|\[.*?\]"
"""
Regular expression: Pattern for matching different pronunciation notations, including:
*w, [w], [ß], [^b], +', and any other notation surrounded by square brackets
"""
clean_location_rx = r"\:"
"""
Regular expression: Pattern for matching the colon symbol in the location information
"""
seperator_rx = r"(((?<=\,|\/).*)| .([A-Z][A-Z][A-Z]:|OU:).*)"
"""
Regular expression: Pattern for matching text separated by a comma or a forward slash,
or a space followed by a three capital letter location identifier with a colon
"""
number_rx = r"(\+\d{1,2}|\s\d\d\s|\d\d\D\s|\d\dx|\d\d[A-Z]$)"
"""
Regular Expression: pattern to match any sequence of digits (0-9) to capture numbers.
"""

tomini_languages = [
    'BALAESANG', 'DAMPELAS', 'TAJIO', 'TAJE', 'PETAPA', 'PENDAU', 'LAUJE',
    'BOU', 'AMPI', 'DONDO', 'TIALO', 'TOTOLI', 'BOANO', 'BUOL'
]

def init_table(language, input_data=pd.read_excel(input_filepath)):
    """
    This function loads the data from the specified Excel file and creates a dataframe with the relevant columns. 
    The function also inserts additional columns with default values.
    
    :param language: The language to be used as the form column
    :param input_data: The input data in the form of a pandas dataframe (default: data from the "input_data\Tomini.xlsx" file)
    :return: The created dataframe
    """
    df = input_data[[
        'EXCLUDE', 'CATEGORY1', 'CATEGORY2', 'CATEGORY3', 'CATEGORY4', 'HOLLE',
        'SIL', 'REID', 'BLUST', 'LIST', 'SWA100', 'GUD200'
    ]].copy()
    df.insert(0, "NUMBER", input_data["NUMBER"])
    df.insert(1, "IDSW", input_data["IDSW"])
    df.insert(2, "ENGLISH", input_data["ENGLISH"])
    df.insert(3, "INDONESIAN", input_data["INDONESIAN"])
    df.insert(4, "FORM", input_data[language])
    df.insert(5, "ORIGINAL_FORM", input_data[language])
    df.insert(6, "LANGUAGE", language)
    df.insert(7, "LOCATION", language)
    df.insert(8, "LONGITUDE", None)
    df.insert(9, "LATITUDE", None)
    df.insert(10, "GENDER", None)
    df.insert(11, "RECORDED", None)
    df.insert(12, "MEDIALINK",
              None)  # Audiodatei Benennung: Sprachkürzel ORT_NUMBER_0
    df.insert(13, "REGISTER", None)
    df.insert(15, "REFERENCE", None)
    df.insert(16, "GRAMMAR_COMMENT", None)
    df.insert(17, "PRONUNCIATION_COMMENT", None)
    df.insert(18, "TRANSLATION_COMMENT", None)
    df.insert(19, "EDITING_COMMENT", None)
    return df

def trim_column(df, column):
    df[column] = df[column].apply(lambda x: str(x).strip())
    return df


def clean(df, column, pattern):
    df[column].replace(to_replace=pattern, value="", regex=True, inplace=True)
    return df

def init_preprocessing(df):
    """
        This function performs pre-processing on the input dataframe. The pre-processing steps include cleaning the 
        "FORM" column by removing specified patterns using the "clean" function, and replacing specific patterns in the
        "FORM" column with new values using the pandas replace method.
        
        :param df: The input dataframe
        :return: The pre-processed dataframe
    """
    df = clean(df, "FORM", r"(p\d+\w|\d+\/(\d{1,}|\w)|p\d)")
    df["FORM"].replace(to_replace="(-)(/)",
                       value="\g<1>$",
                       regex=True,
                       inplace=True)
    df["FORM"].replace(to_replace="((\+)([A-Z][A-Z][A-Z])(:))",
                       value="\g<2>\g<3>, \g<3>\g<4>",
                       regex=True,
                       inplace=True)
    df["FORM"].replace(to_replace="(  (?!RDP)([A-Z]{3}:))",
                       value=" ,\g<2>",
                       regex=True,
                       inplace=True)
    df["FORM"].replace(to_replace="((_)(\:))",
                       value="\g<3>",
                       regex=True,
                       inplace=True)
    df["FORM"].replace(to_replace="(\([^)]*)(,)([^)]*\))",
                       value="\g<1>;\g<3>",
                       regex=True,
                       inplace=True)
    return df

def rename_location(df, pattern):
    """
    Rename the 'LOCATION' column based on the values matching the provided pattern in the 'FORM' column.
    The matching values will be extracted and assigned to the 'LOCATION' column.
    The matching values will also be removed from the 'FORM' column.
    
    Parameters:
    df (DataFrame): The dataframe to perform the operation on.
    pattern (str): The pattern to use to extract the values from the 'FORM' column.
    
    Returns:
    DataFrame: The input dataframe with the updated 'LOCATION' and 'FORM' columns.
    """
    df["LOCATION"] = pd.Series(df["FORM"]).str.extract(pattern).values
    df = clean(df, "FORM", "(^([A-Z]{3}|OU):)")
    return df

def add_second_location(df, pattern):
    """
    Extract second location from the 'FORM' column and add it as a new row in the dataframe.
    Parameters:
    df (pd.DataFrame): The input dataframe.
    pattern (str): The pattern used to extract the second location.
    Returns:
    df (pd.DataFrame): The updated dataframe with the second location added.
    """
    new = df.merge(df["FORM"].str.extractall(location2_rx).reset_index(
        level=-1, drop=True),
                   left_index=True,
                   right_index=True)
    new.drop(columns="LOCATION", inplace=True)
    new.rename({0: "LOCATION"}, inplace=True, axis="columns")
    df = pd.concat([df, new], ignore_index=True)
    clean(df, "FORM", pattern)
    clean(df, "LOCATION", "\+")
    return df

def seperate_lexems(df):
    """
    This function separates lexems in the "FORM" column of the dataframe by splitting them based on ',' or '/'
    and stacks the resulting series into a single dataframe.
    The split forms are then joined back to the original dataframe and the original "FORM" column is temporarily
    stored and dropped, with the temporary column being renamed to "FORM".
    The index of the dataframe is then reset.
    
    :param df: The input dataframe
    :return: The updated dataframe with separated lexems in the "FORM" column
    """
    s = df["FORM"].str.split(',|/').apply(pd.Series, 1).stack()
    s.index = s.index.droplevel(-1)
    s.name = "FORM"
    del df["FORM"]
    df = df.join(s)
    df.insert(4, "FORM_temp", df["FORM"])
    df.drop('FORM', axis=1, inplace=True)
    df.rename(columns={'FORM_temp': 'FORM'}, inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df

def add_recording(df, pattern):
    """
    This function adds the recording information to the dataframe by checking if the `FORM` column contains
    the given `pattern` and adding the result to the `RECORDED` column.
    The `FORM` column is then cleaned to remove the `pattern` from its values.
    
    :param df: The input dataframe
    :param pattern: The pattern to search for in the `FORM` column
    :return: The `RECORDED` column
    """
    list = df["FORM"].str.contains(re.compile(pattern)).values
    df["RECORDED"] = np.logical_not(list)
    clean(df, "FORM", pattern)
    return df["RECORDED"]

def add_gender(df, pattern):
    """
    This function adds the gender information to the dataframe by extracting the gender information from the FORM column
    using the given pattern and adding the result to the GENDER column.
    The FORM column is then cleaned to remove the pattern from its values.
    The extracted gender information is then replaced with either 'feminine' or 'masculine'.
    
    :param df: The input dataframe
    :param pattern: The pattern to search for in the `FORM` column
    :return: The `GENDER` column
    """
    df["GENDER"] = pd.Series(df["FORM"]).str.extract(pattern).values
    clean(df, "FORM", pattern)
    df["GENDER"].replace(to_replace="(per)", value="feminine", inplace=True)
    df["GENDER"].replace(to_replace="(l-l)", value="masculine", inplace=True)
    return df["GENDER"]

def add_register(df, pattern):
    """
    This function adds the register information to the dataframe by extracting values from the "FORM" column that match the `pattern` and adding the result to the "REGISTER" column. The "FORM" column is then cleaned to remove the `pattern` from its values.
    
    :param df: The input dataframe
    :param pattern: The pattern to search for in the "FORM" column
    :return: The "REGISTER" column
    """
    df["REGISTER"] = pd.Series(df["FORM"]).str.findall(pattern)
    clean(df, "FORM", pattern)
    df["REGISTER"] = df["REGISTER"].astype(str)
    clean(df, "REGISTER", group_rx)
    return df["REGISTER"]

def add_grammar_comment(df, pattern):
    """
    This function adds the grammar comment information to the dataframe by checking if the `FORM` column contains
    the given `pattern` and adding the result to the `GRAMMAR_COMMENT` column.
    The `FORM` column is then cleaned to remove the `pattern` from its values.
    Additionally, the function replaces all occurrences of '$' in the `GRAMMAR_COMMENT` column with '/'.
    
    :param df: The input dataframe
    :param pattern: The pattern to search for in the `FORM` column
    :return: The `GRAMMAR_COMMENT` column
    """
    df["GRAMMAR_COMMENT"] = pd.Series(df["FORM"]).str.extract(pattern).values
    df["GRAMMAR_COMMENT"].replace(to_replace=r"\$",
                                  value="/",
                                  regex=True,
                                  inplace=True)
    clean(df, "FORM", pattern)
    return df["GRAMMAR_COMMENT"]

def add_translation_comment(df, pattern):
    """
    Adds the contents of the translation comment to the "TRANSLATION_COMMENT" column of the dataframe. The contents are extracted from the "FORM" column using the specified pattern, and the matched values are removed from the "FORM" column.

    :param df: The input dataframe
    :param pattern: The bpattern used to extract the translation comment from the "FORM" column
    :return: The "TRANSLATION_COMMENT" column of the dataframe
    """
    df["TRANSLATION_COMMENT"] = pd.Series(df["FORM"]).str.findall(pattern)
    clean(df, "FORM", pattern)
    df["TRANSLATION_COMMENT"] = df["TRANSLATION_COMMENT"].astype(str)
    clean(df, "TRANSLATION_COMMENT", group_rx)
    return df["TRANSLATION_COMMENT"]

def add_pronunciation_comment(df, pattern):
    """Extracts pronunciation comments from the "FORM" column and adds them to a new column "PRONUNCIATION_COMMENT".

    :param df: Dataframe
    :param pattern: Regex pattern
    :return: "PRONUNCIATION_COMMENT" column
    """
    df["PRONUNCIATION_COMMENT"] = pd.Series(df["FORM"]).str.findall(pattern)
    clean(df, "FORM", pattern)
    df["PRONUNCIATION_COMMENT"] = df["PRONUNCIATION_COMMENT"].astype(str)
    clean(df, "PRONUNCIATION_COMMENT", group_rx)
    return df["PRONUNCIATION_COMMENT"]

def add_editing_comment(df, pattern):
    """
    Extract editing comments from the "FORM" column and add it as a new column in the dataframe.
    
    :param df: The input dataframe
    :param pattern: The pattern used to extract the editing comment from the "FORM" column
    :return: The "EDITING_COMMENT" column of the dataframe
    """
    df["EDITING_COMMENT"] = pd.Series(df["FORM"]).str.findall(pattern)
    clean(df, "FORM", pattern)
    df["EDITING_COMMENT"] = df["EDITING_COMMENT"].astype(str)
    clean(df, "EDITING_COMMENT", group_rx)
    return df["EDITING_COMMENT"]

def add_reference(df, pattern):
    """
    Extract reference information from the "FORM" column and add it as a new column in the dataframe.
    
    :param df: The input dataframe
    :param pattern: The pattern used to extract the reference information
    :return: The updated dataframe
    """
    df["REFERENCE"] = pd.Series(df["FORM"]).str.extract(pattern).values
    clean(df, "FORM", pattern)
    df["REFERENCE"] = df["REFERENCE"].astype(str)
    clean(df, "REFERENCE", group_rx)
    return df["REFERENCE"]

def add_missing_references(df):
    """
    Finds missing references in the 'FORM' column and updates the values in the column with the found references.
    The function first creates a filtered dataframe containing the missing references, which can be identified by
    the presence of the reference2_rx pattern in the 'REFERENCE' column and an empty value in the 'FORM' column.
    The function then updates the 'FORM' column of the filtered dataframe with the values from the 'REFERENCE' column.
    Finally, the function updates the 'FORM' column of the input dataframe with the values from the filtered dataframe.

    :param df: The input dataframe
    :return: The updated dataframe with missing references added to the 'FORM' column.
    """
    filter_df = df.loc[(df["REFERENCE"].str.contains(reference2_rx))
                       & (df["FORM"] == '')].copy()
    filter_df["FORM"] = filter_df.loc[:, "REFERENCE"]
    filter_df["FORM"] = filter_df.loc[:, "FORM"].str.strip("=|+")
    filter_df.replace(dict(zip(df["IDSW"], df['FORM'])), inplace=True)
    df.update(filter_df.FORM)
    return df

def add_location_data(df, language):
    """
    Add location data to the dataframe.
    This function adds the latitude, longitude, and language information to the input dataframe. If the "LOCATION" column is missing, it is filled with the provided language. The location information is then updated to the full name of the location. The latitude and longitude information is also added to the dataframe using dictionaries with the location names as keys and the corresponding latitudes and longitudes as values.

    :param df: The input dataframe
    :param language: The language to use if the "LOCATION" column is missing
    :return: The updated dataframe with location information
    """
    df["LOCATION"].fillna(language, inplace=True)
    df["LOCATION"].replace(default_hamlet, inplace=True)
    df.replace(abbreviation, inplace=True)
    df["LATITUDE"] = df["LOCATION"].copy()
    df["LATITUDE"].replace(latitude, inplace=True)
    df["LONGITUDE"] = df["LOCATION"].copy()
    df["LONGITUDE"].replace(longitude, inplace=True)
    df["LANGUAGE"] = df["LANGUAGE"].str.capitalize()
    return df

def init_postprocessing(df):
    """
    Perform post-processing on the input dataframe. The post-processing steps include:
    - Replacing values in the "LANGUAGE" column using a dictionary.
    - Replacing values in the "FORM" column using a dictionary and regular expressions.
    - Trimming the "FORM" column.
    - Dropping rows with null values in either the "FORM" or "ORIGINAL_FORM" columns (keeping at least two non-null values).
    - Dropping rows with null values in the "FORM" column.
    
    :param df: The input dataframe.
    :return: The post-processed dataframe.
    """
    df["LANGUAGE"].replace({"Petapa": "Taje"}, inplace=True)
    df["LANGUAGE"].replace({"Bou": "Lauje"}, inplace=True)
    df["FORM"].replace(to_replace="([A-Z]{3}$)",
                       value="+\g<1>",
                       regex=True,
                       inplace=True)
    df["FORM"].replace(to_replace="(_\W)", value="", regex=True, inplace=True)
    df["FORM"].replace(to_replace="(#$)", value="", regex=True, inplace=True)
    trim_column(df, "FORM")
    df.dropna(subset=["FORM", "ORIGINAL_FORM"], thresh=2, inplace=True)
    df.dropna(subset=["FORM"], inplace=True)
    return df

def sort(df, column):
    """
    Sort the input dataframe by a specified column.
    Parameters:
    df (pd.DataFrame): The input dataframe.
    column (str): The name of the column to sort the dataframe by.
    Returns:
    df (pd.DataFrame): The sorted dataframe.
    """
    df.sort_values(by=[column], inplace=True)
    return df

def sort_by_number(df):
    """
    Sort the input dataframe by the 'NUMBER' column.
    Parameters:
    df (pd.DataFrame): The input dataframe.
    Returns:
    df (pd.DataFrame): The sorted dataframe.
    """
    df.sort_values(by=["NUMBER"], inplace=True)
    return df

def sort_by_idsw(df):
    """
    Sort the input dataframe by the 'IDSW' column.
    Parameters:
    df (pd.DataFrame): The input dataframe.
    Returns:
    df (pd.DataFrame): The sorted dataframe.
    """
    df.sort_values(by=["IDSW"], inplace=True)
    return df

def search(df, word):
    """
    Search the input dataframe for rows containing a specified word in the 'FORM' column.
    
    Parameters:
    df (pd.DataFrame): The input dataframe.
    word (str): The word to search for in the 'FORM' column.
    
    Returns:
    result (pd.DataFrame): The rows from the input dataframe containing the specified word in the 'FORM' column.
    """
    result = df[df['FORM'].str.contains(word)]
    return result

def search_id(df, id):
    """
    Search the input dataframe for a row with a specified ID.
    
    Parameters:
    df (pd.DataFrame): The input dataframe.
    id (int): The ID to search for in the dataframe.
    
    Returns:
    result (pd.Series): The row from the input dataframe with the specified ID.
    """
    df.sort_values(by=["NUMBER"], inplace=True)
    return df.loc[
        id,
        ["IDSW", "FORM", "ORIGINAL_FORM", "LANGUAGE", "TRANSLATION_COMMENT"]]

def search_idsw(df, idsw):
    """
    Search the input dataframe for rows containing a specified IDSW.
    
    Parameters:
    df (pd.DataFrame): The input dataframe.
    idsw (str): The IDSW to search for in the 'IDW' column.
    
    Returns:
    result (pd.DataFrame): The rows from the input dataframe containing the specified IDSW in the 'IDW' column.
    """
    result = df[df['IDSW'].str.contains(idsw, na=False)]
    return result

def create_complete_df():
    """Create a complete dataframe by processing all the tables for each language.

    :return: The concatenated dataframe of all the processed tables.
    """
    print("Processing tables for each language...")
    for language in tqdm(tomini_languages):
        df = init_table(language)
        init_preprocessing(df)
        df = seperate_lexems(df)
        df = add_features(df)
        add_comments(df)
        add_location_data(df, language)
        init_postprocessing(df)
        df_list.append(df)
    print("Tables processed successfully")
    return pd.concat(df_list)

def add_features(df):
    """
    Extract and add features to the dataframe.

    :param df: The input dataframe.
    :return: The dataframe with extracted and added features.
    """
    add_reference(df, reference_rx)
    df = rename_location(df, location1_rx)
    df = add_second_location(df, location2_rx)
    add_register(df, register_rx)
    add_recording(df, recorded_rx)
    add_missing_references(df)
    df = clean(df, "FORM", location1_rx)
    df = clean(df, "FORM", number_rx)
    return df

def add_comments(df):
    """
    Extract and add comments to the dataframe.

    :param df: The input dataframe.
    :return: The dataframe with extracted and added comments.
    """
    add_editing_comment(df, editing_comment_rx)
    add_gender(df, gender_rx)
    add_translation_comment(df, translation_comment_rx)
    add_pronunciation_comment(df, pronunciation_rx)
    add_grammar_comment(df, grammar_rx)
    df = clean(df, "LOCATION", r":")
    df = clean(df, "GRAMMAR_COMMENT", r"(\[.|.\])")
    return df

def create_complete_csv(output_path):
    """
    Write the concatenated dataframe to a csv file.
    
    :param output_path: The file path where the csv file will be saved.
    :return: None
    """
    print("Creating Tomini_Complete CSV...")
    tomini_all = create_complete_df()
    tomini_all.to_csv(output_path, encoding='utf-8-sig')
    print(f"CSV file created at: {output_path}")

def create_split_csv(output_path):
    """
    Write a separate csv file for each language in the `tomini_languages` list.

    :param output_path: The directory where the csv files will be stored.
    :return: None
    """
    print("Saving all languages into different csv files")
    for language in tqdm(tomini_languages):
        df = init_table(language)
        init_preprocessing(df)
        df = seperate_lexems(df)
        df = add_features(df)
        add_comments(df)
        add_location_data(df, language)
        init_postprocessing(df)
        df.to_csv(f"{output_path}/{language}.csv",
                  encoding='utf-8-sig',
                  index=False)
    print("Split CSV files successfully created!")

def create_split_xlsx(output_path):
    """
    Create an XLSX file that contains the data from the tomini_languages
    list. Each language will have its own sheet within the XLSX file.

    :param output_path: The file path where the XLSX file will be saved.
    :return: None
    """
    writer = pd.ExcelWriter(output_path,
                            engine='xlsxwriter',
                            engine_kwargs={'options': options})
    print("Saving all languages into different worksheets to {}".format(
        output_path))
    for language in tqdm(tomini_languages):
        df = init_table(language)
        init_preprocessing(df)
        df = seperate_lexems(df)
        add_features(df)
        add_comments(df)
        add_location_data(df, language)
        init_postprocessing(df)
        df.to_excel(writer,
                    sheet_name='{}.xlsx'.format(language),
                    encoding="utf-8",
                    na_rep="nan",
                    verbose=False,
                    index=True)
    writer.close()
    print("XLSX file saved successfully!")


def create_complete_xlsx(output_path):
    """
    Writes the complete dataframe to an xlsx file.
    :param output_path: The path where the xlsx file will be stored.
    """
    writer = pd.ExcelWriter(output_path,
                            engine='xlsxwriter',
                            engine_kwargs={'options': options})
    print("Creating Tomini_Complete XLSX Table...")
    df_total = create_complete_df()
    for i, row in tqdm(df_total.iterrows(), total=df_total.shape[0]):
        df_total.iloc[[i]].to_excel(writer,
                                    index=False,
                                    header=False,
                                    startrow=i)
    writer.save()
    print(f"Tomini_Complete Table successfully created at: {output_path}")


if __name__ == "__main__":
    create_split_csv("output/csv")
    create_complete_csv("output/csv/Tomini_complete.csv")
    create_complete_xlsx("output/excel/Tomini_complete.xlsx")
    create_split_xlsx("output/excel/Tomini_split.xlsx")
