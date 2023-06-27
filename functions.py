import re


def replace_text(match, parameters):
    """

    Parameters
    ----------
    match : re.Match
        Match from the find_and_replace_text method.
    parameters : pandas.Series
        Series with parameter values to be replaced, the parameter name has to be in the index name.
    """
    for name, value in parameters.items():
        if match.group(1).casefold() == '@' + name.casefold():  # find, case-insensitive
            return match.group(1)[1:] + '=' + str(value)


def find_and_replace_text(path_file, pattern, parameters):
    """Find and replace text within txt file.

    Find text within a txt file which matches a certain pattern and replace it.

    Parameters
    ----------
    path_file : str
        Path of the .txt file
    pattern : str
        Pattern which marks parameters to be replaced
    parameters : pandas.Series
        Series with parameter values to be replaced, the parameter name has to be in the index name.
    """

    # read file
    with open(path_file, 'r') as file:
        text = file.read()

        # find, then replace text values
        text = re.sub(pattern=pattern, repl=lambda match: replace_text(match, parameters), string=text)

    # overwrite file
    with open(path_file, 'w') as file:
        file.write(text)
