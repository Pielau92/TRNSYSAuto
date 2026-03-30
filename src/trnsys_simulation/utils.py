import os
import re


def parent_dir(path: str, levels: int = 1) -> str:
    """Get parent directory.

    :param str path: starting path
    :param int levels: specifies how many directory levels to move upwards from the starting path
    :return: parent directory path
    """

    for _ in range(levels):
        path = os.path.dirname(path)
    return path


def find_and_replace(path_file: str, pattern: str, replacement: str) -> None:
    """Find string in text and replace by another string.

    Reads text from a .txt file, searches for a defined string, replaces it by another string and overwrites the .txt
    file.

    :param str path_file: path to txt file
    :param str pattern: pattern which marks the text to be replaced
    :param str replacement: text to be inserted instead
    """

    with open(path_file, 'r') as file:
        text = file.read()  # read file
        new_text = re.sub(pattern=pattern, repl=replacement, string=text)
    with open(path_file, 'w') as file:
        file.write(new_text)  # overwrite file


def replace_parameter_values(path_file: str, parameters: dict, mark: bool = False) -> None:
    """Find and replace parameter values within a .txt file.

    Finds parameter values within a .txt file, replaces them and overwrites the .txt file. For this, the following must
    apply:
    -   A parameter means in this case a name/value pair inside a txt file - specifically: a word followed by "=" and a
        number (there may be any amount of white spaces or tabs between the word and "=" and between "=" and the number)
    -   The replacement values have to be passed in "parameters", where the row indices must match the name of the
        parameter whose number value is to be replaced

    Example:
        # .txt file containing "Parameter1=1; Parameter2=2; @Parameter3=3."
        path_file = 'C:\\Users\\JohnDoe\\Desktop\\text.txt'

        parameters = pd.Series({'Parameter1': 4, 'Parameter2': 5, 'Parameter3': 6})

        replace_parameter_value(path_file, parameters)

        # .txt file now shows "Parameter1=4; Parameter2=5; Parameter3=6."

    :param str path_file: path of txt file
    :param dict parameters: dict with parameters as key value pairs
    """

    def replacer(match):
        parameter = match.group(1)  # parameter name
        initial_value = match.group(2).strip()
        comment = match.group(3)  # miscellaneous characters after the number value (typically comments)

        if comment is None:
            comment = ''

        if parameter in parameters.keys():
            new_value = str(parameters[parameter])
        else:
            return match.group(0)  # return unchanged text

        if initial_value == new_value:
            return match.group(0)  # return unchanged text

        if mark:  # mark changed lines with comment
            comment += f' ! *parameter changed from {initial_value} to {new_value}*'  # add remark

        return f"{parameter} = {new_value} {comment}"  # replace, if parameter name matches

    with open(path_file, 'r') as file:
        text = file.read()

    pattern = re.compile(
        r'^(\w*)'  # word at beginning of line
        + r'[\s\t]*=[\s\t]*'  # equal sign (=), with any number of white spaces/tabs before and after
        + r'([^\n!]*)'  # any number of characters, that are neither an exclamation mark nor the end of the line
        + r'(?:(!.*))?'  # (optional) exclamation mark followed by any number of characters
        + r'$',  # end of line
        re.MULTILINE
    )

    new_text = pattern.sub(replacer, text)

    # overwrite file
    with open(path_file, 'w') as file:
        file.write(new_text)


def delete_files(paths: list[str]) -> None:
    """Delete multiple files.

    :param paths: list of paths to files to be deleted
    """

    for path in paths:
        try:
            os.remove(path)
        except FileNotFoundError:
            pass
