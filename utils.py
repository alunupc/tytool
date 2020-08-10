from convert import doc_to_pdf


def process_param(param):
    if param.endswith(","):
        param = param.rstrip(",")
    if '-' in param:
        if int(param.split('-')[0]) < int(param.split('-')[1]):
            param = list(range(int(param.split('-')[0]), int(param.split('-')[1]) + 1))
            return param
        else:
            return []
    else:
        if ',' in param:
            param = list(map(int, param.split(',')))
            param.sort()
            return param
        elif isinstance(param, str) and param.isdigit():
            return [int(param)]
        else:
            return []


def process_file(file_path):
    suffix = file_path.split(".")[-1]
    if suffix.lower() == "doc" or suffix.lower() == "docx":
        doc_to_pdf(file_path, "." + suffix)
