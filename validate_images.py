import re


def get_image_filenames(message_filepath):
    message = open(message_filepath).read()
    return re.findall('<img [^>]*src="([^"]+)', message)


