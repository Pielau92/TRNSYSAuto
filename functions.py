import re


def replace_param(match, parameters):
    """

    Parameters
    ----------
    match :
    parameters :

    Returns
    -------

    """
    for name, value in parameters.items():
        if match.group(1).casefold() == '@' + name.casefold():  # find, case-insensitive
            return match.group(1)[1:] + '=' + str(value)


def find_and_replace_param(path_file, pattern, parameters):
    """

    Parameters
    ----------
    path_file : path of the .txt file
    pattern :
    parameters :

    Returns
    -------

    """

    # read file
    with open(path_file, 'r') as file:
        text = file.read()

        # find, then replace parameter values
        text = re.sub(pattern=pattern, repl=lambda match: replace_param(match, parameters), string=text)

    # overwrite file
    with open(path_file, 'w') as file:
        file.write(text)
