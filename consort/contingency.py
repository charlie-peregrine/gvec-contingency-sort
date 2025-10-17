# contingency.py, Charlie Jordan, 10/17/2025

import re


big_regex = re.compile(
    r"\/\* Contingency '(.*)'\n"
    r"\/\* StartDate: ([\d/]*); StopDate: ([\d/]*);\n"
    r"\/\* Submitter: (.*); NERCCategory: (.*); ERCOTCategory: (.*);\n"
    r"\s*CONTINGENCY '(.+)'.*?(?:\/\* (.*))?\n"
    r"((?:\s*[\S\s]+?\n)*?)\s*END\n",
    re.MULTILINE
)


class Contingency():
    def __init__(self, full_str, source_file=None):
        self.full_str = full_str
        self.trimmed_full_str = re.sub(r'[ \t]+\/\*.*\n', '\n', full_str)
        self.source_file = source_file
        m = re.match(big_regex, full_str)
        if m is None:
            raise ValueError
        self.name = m[1]
        self.start_date = m[2]
        self.stop_date = m[3]
        self.submitter = m[4]
        self.nerc_cat = m[5]
        self._set_contingency_group()
        self.ercot_cat = m[6]
        self.id = m[7]
        self.id_comment = m[8] if m[8] is not None else ''
        self.lines = []
        for line in m[9].split('\n'):
            if line:
                stripped = line.strip()
                split = re.split(r'\s*\/\*\s*', stripped, maxsplit=1)
                if len(split) == 1:
                    split.append('')
                self.lines.append(tuple(split))
        self.lines_str = '\n'.join(x[0] for x in self.lines)
        self.duplicates = []
    
    def _set_contingency_group(self):
        if "EE1" in self.nerc_cat:
            self.nerc_group = "EXTREME EVENT / DOUBLE CIRCUIT"
        elif "EE2" in self.nerc_cat:
            self.nerc_group = "EXTREME EVENT / LOCAL AREA"
        elif "EE3" in self.nerc_cat:
            self.nerc_group = "EXTREME EVENT / WIDE AREA"
        elif "P1.1" in self.nerc_cat:
            self.nerc_group = "SINGLE GENERATOR OUTAGE"
        elif "P1.2" in self.nerc_cat:
            self.nerc_group = "SINGLE LINE OUTAGE"
        elif "P1.3" in self.nerc_cat:
            self.nerc_group = "SINGLE TRANSFORMER OUTAGE"
        elif "P1.4" in self.nerc_cat:
            self.nerc_group = "SINGLE SHUNT OUTAGE"
        elif "P2." in self.nerc_cat:
            self.nerc_group = "FAULT / PROTECTION EVENT"
        elif "P4." in self.nerc_cat:
            self.nerc_group = "FAULT / PROTECTION EVENT"
        elif "P5." in self.nerc_cat:
            self.nerc_group = "FAULT / PROTECTION EVENT"
        elif "P6." in self.nerc_cat:
            self.nerc_group = "SINGLE POLE OF A DC LINE"
        elif "P7." in self.nerc_cat:
            self.nerc_group = "DOUBLE LINE OUTAGE"
        else:
            print("missing con group!", self.nerc_cat)
            self.nerc_group = "UNSPECIFIED"
    
    def change_id(self, new_id: str):
        old_id = self.id
        self.full_str = self.full_str.replace(old_id, new_id)
        self.trimmed_full_str = self.trimmed_full_str.replace(old_id, new_id)
        for i in range(len(self.lines)):
            if old_id in self.lines[i]:
                self.lines[i] = self.lines[i].replace(old_id, new_id)
        self.lines_str.replace(old_id, new_id)
        self.id = new_id

    def make_csv_line_dict(self):
        return {
            'CONTINGENCY': self.id,
            'CONTINGENCY TYPE': f"NERC {self.nerc_cat.replace('/', ' / ')}",
            'CONTINGENCY GROUP': self.nerc_group,
            'CONTINGENCY DESCRIPTION': self.trimmed_full_str.rstrip(),
            'DATE START': self.start_date,
            'DATE END': self.stop_date,
            'DUPLICATES' : '\n'.join(
                (f"{c.id} - {c.nerc_cat}" for c in self.duplicates)
            )
        }
    
    def __hash__(self) -> int:
        return hash(self.full_str)
    
    def __eq__(self, other: object) -> bool:
        return hash(self) == hash(other)

