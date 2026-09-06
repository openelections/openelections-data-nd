# Parse the 2026 primary election results into county- and precinct-level CSVs.
#
# Source: ../openelections-sources-nd/2026_Primary_Election.csv, a statewide
# export from resultsnd.sos.nd.gov (the ENR site's "download full data" file)
# with one row per contest/location/candidate. Each location appears twice:
# as a county row ("Traill County", Location ID = 5-digit FIPS) and as
# precinct rows ("Traill County - 2010", Location ID = precinct number).
#
# Conventions match the existing files in this repo (2018/2020):
#   - office names: State Representative -> State House, State Senator ->
#     State Senate, Representative in Congress -> U.S. House
#   - district: legislative district number without leading zeros
#   - party: DEM/REP from the Party ID column; blank for nonpartisan
#   - county: without the " County" suffix, Mc-names in their proper casing
#   - candidate votes of "undefined" become 0
import csv
import os
import re
from collections import defaultdict

SOURCE = os.path.join(os.path.dirname(__file__), '../openelections-sources-nd/2026_Primary_Election.csv')
COUNTY_OUTPUT = os.path.join(os.path.dirname(__file__), '2026/20260609__nd__primary__county.csv')
PRECINCT_OUTPUT = os.path.join(os.path.dirname(__file__), '2026/20260609__nd__primary__precinct.csv')

COUNTY_FIXUPS = {
    'Mchenry': 'McHenry',
    'Mcintosh': 'McIntosh',
    'Mckenzie': 'McKenzie',
    'Mclean': 'McLean',
}

# Office names standardized to match the existing files.
OFFICE_RENAMES = {
    'State Representative': 'State House',
    'State Representative Unexpired 2-Year Term': 'State House Unexpired 2-Year Term',
    'Representative in Congress': 'U.S. House',
}

PARTY_SUFFIX = re.compile(r' (Democratic-NPL|Republican)$')


def clean_county(location_name):
    name = location_name[:-len(' County')]
    return COUNTY_FIXUPS.get(name, name)


def clean_name(name):
    # collapse runs of whitespace, e.g. "Shelly   Bruse " -> "Shelly Bruse"
    return re.sub(r'\s+', ' ', name.strip())


def parse_contest(contest_name, party_id):
    """Return (office, district, party) for a contest name."""
    party = party_id.strip() if party_id.strip() in ('DEM', 'REP') else ''
    office = PARTY_SUFFIX.sub('', clean_name(contest_name))
    district = ''
    m = re.search(r'\bState (Senator|Representative)(.*?) District (\d+)$', office)
    if m:
        office = 'State ' + ('Senate' if m.group(1) == 'Senator' else 'House') + m.group(2)
        district = str(int(m.group(3)))
    office = OFFICE_RENAMES.get(office, office)
    if office == 'U.S. House':
        district = '1'
    return office.strip(), district, party


def parse_location(location_name):
    """Split "Traill County - 2010" into ("Traill", "2010"); county rows
    ("Traill County") have no precinct."""
    if ' - ' not in location_name:
        return clean_county(location_name), ''
    county, precinct = location_name.split(' - ', 1)
    return clean_county(county), clean_name(precinct)


def main():
    os.makedirs(os.path.join(os.path.dirname(os.path.abspath(__file__)), '2026'), exist_ok=True)

    rows = []
    with open(SOURCE, newline='', encoding='utf-8-sig') as f:
        for r in csv.DictReader(f):
            county, precinct = parse_location(r['Location Name'])
            votes = r['Candidate Votes'].strip()
            votes = 0 if votes == 'undefined' else int(votes)
            office, district, party = parse_contest(r['Contest Name'], r['Party ID'])
            rows.append((county, precinct, office, district, party,
                         clean_name(r['Candidate Name']), votes))

    # Aggregate on the output key. This merges the one case where two contest
    # IDs share a name in the same county (Stutsman's Jamestown school board
    # seats, whose "write-in" rows would otherwise be duplicates).
    def write(path, header, key, keep=lambda row: True):
        totals = defaultdict(int)
        for row in rows:
            if keep(row):
                totals[key(row)] += row[-1]
        with open(path, 'wt', newline='') as f:
            w = csv.writer(f)
            w.writerow(header)
            for k in sorted(totals):
                w.writerow([*k, totals[k]])
        print(f'wrote {len(totals)} rows to {path}')

    write(COUNTY_OUTPUT,
          ['county', 'office', 'district', 'party', 'candidate', 'votes'],
          lambda row: (row[0], *row[2:6]),
          keep=lambda row: row[1] == '')
    write(PRECINCT_OUTPUT,
          ['county', 'precinct', 'office', 'district', 'party', 'candidate', 'votes'],
          lambda row: row[:6],
          keep=lambda row: row[1] != '')


if __name__ == '__main__':
    main()