import re


def find_and_replace_param(path_file, pattern, parameters):
    """

    Parameters
    ----------
    path_file :
    pattern :
    parameters :

    Returns
    -------

    """
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
            if match.group(1).casefold() == '@' + name.casefold():
                return match.group(1)[1:] + '=' + str(value)

    with open(path_file, 'r') as file:
        text = file.read()  # read file

        # find, then replace parameter values
        text = re.sub(pattern=pattern, repl=lambda match: replace_param(match, parameters), string=text)

    with open(path_file, 'w') as file:
        file.write(text)  # overwrite file
