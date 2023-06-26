import json

COMMA = ","
EMPTY_STRING = ""
NEWLINE = "\n"

class ClaimsDataHeader:
    def __init__(self, line: str, ) -> None:
        self.fields = line.split(COMMA)
        for x in range(0,len(self.fields)):
            self.fields[x] = self.fields[x].strip()

class ClaimsData:
    def __init__(self, data_header: ClaimsDataHeader, line: str) -> None:
        self.data_header = data_header
        self.data = line.split(COMMA)
        for x in range(0,len(self.data)):
            self.data[x] = self.data[x].strip()

def get_field_map(filename: str):
    data = None
    with open(filename) as f:
        # Load the contents of the file into a Python object
        data = json.load(f)

    return data

def read_claims_data(tran: str):
    result = []
    with open(tran + '_test.csv', 'r') as f:
        first = True
        data_header = None
        # Loop over each line in the file
        for line in f:
            line = line.replace(NEWLINE, EMPTY_STRING)
            if first:
                data_header = ClaimsDataHeader(line)
                first = False
            else:
                claim = ClaimsData(data_header, line)
                result.append(claim)

    return result

def write_file(file: str, data):
    _write_file_data(file, data, "w")

def append_file(file: str, data):
    _write_file_data(file, data, "a")

def _write_file_data(file: str, data, method: str):
    f = open(file,method)
    f.write(data)
    f.close()

def pad_char(l: int, c: str):
    result = ""
    for x in range(0,l):
        result = result + c

    return result