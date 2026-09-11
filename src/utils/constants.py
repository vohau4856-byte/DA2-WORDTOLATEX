

PANDOC_ARGS = [
    '--standalone',
    '--wrap=preserve',
    '--columns=999',
    '-f', 'docx',       
    '-t', 'latex'
]

MEDIA_DIR_NAME = "extracted_media"