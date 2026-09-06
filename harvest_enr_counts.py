# Harvest county-level vote totals as DISPLAYED on the ENR county pages.
#
# For most elections the per-contest workbooks' TOTALS rows equal the site's
# final tally (verified against ResultsExport.aspx's statewide CSV and, for
# 2004, the certified canvass). The 2006 general is the exception: its
# workbooks carry only ~84% of the final votes while the page display matches
# the certified totals (Pomeroy 142,934). This script harvests the
# display-results boxes on each county page into the manifest as
# 'county_page_votes': {county: {candidate: votes}}, which parse_enr_exports
# then uses for the county-level rows of that election.
#
# Usage: python3 harvest_enr_counts.py <eid> <tag>
import json
import os
import re

from scrape_enr_exports import COUNTIES, OUT_DIR, election_url, session

# the county pages carry no county name; cty=01..53 map to the alphabetical
# county list (verified against the workbooks' alphabetical county sheets)
COUNTIES_ALPHA = [
    'Adams', 'Barnes', 'Benson', 'Billings', 'Bottineau', 'Bowman', 'Burke',
    'Burleigh', 'Cass', 'Cavalier', 'Dickey', 'Divide', 'Dunn', 'Eddy',
    'Emmons', 'Foster', 'Golden Valley', 'Grand Forks', 'Grant', 'Griggs',
    'Hettinger', 'Kidder', 'LaMoure', 'Logan', 'McHenry', 'McIntosh',
    'McKenzie', 'McLean', 'Mercer', 'Morton', 'Mountrail', 'Nelson',
    'Oliver', 'Pembina', 'Pierce', 'Ramsey', 'Ransom', 'Renville',
    'Richland', 'Rolette', 'Sargent', 'Sheridan', 'Sioux', 'Slope', 'Stark',
    'Steele', 'Stutsman', 'Towner', 'Traill', 'Walsh', 'Ward', 'Wells',
    'Williams',
]

BOX_D = re.compile(r'display-results-box-d[^>]*>\s*<h2>(.*?)</h2>', re.S)
BOX_G = re.compile(r'display-results-box-g" ?>([\d,]+)</h2>')


def parse_counts(html):
    """Return {contest_id: (title, {candidate: votes})} for one county page."""
    out = {}
    for block in html.split('wrapper-inside wrapper-border')[1:]:
        title_m = re.search(r'<h1>(.*?)</h1>', block, re.S)
        chk_m = re.search(r'id="chk-(\d+)"', block)
        if not (title_m and chk_m):
            continue
        title = re.sub(r'\s+', ' ', re.sub(r'<[^>]+>', ' ', title_m.group(1))).strip()
        # candidate h2s (box-d) pair up with the vote h2s (box-g) in order;
        # the Total Votes row uses box-total, not box-g
        candidates = [re.sub(r'\s+', ' ', re.sub(r'<[^>]+>', ' ', c)).strip()
                      for c in BOX_D.findall(block)]
        candidates = [c for c in candidates if c and c != 'Total Votes']
        votes = [int(v.replace(',', '')) for v in BOX_G.findall(block)]
        if len(votes) != len(candidates):
            raise ValueError(f'{title}: {len(candidates)} candidates vs {len(votes)} votes')
        out[chk_m.group(1)] = (title, dict(zip(candidates, votes)))
    return out


def main(eid, tag):
    election_dir = os.path.join(OUT_DIR, tag)
    manifest_path = os.path.join(election_dir, 'manifest.json')
    with open(manifest_path) as f:
        manifest = json.load(f)

    harvested = {}
    for n in COUNTIES:
        url = election_url(eid, 'ResultsSW.aspx',
                           {'type': 'CTYALL', 'cty': f'{n:02d}', 'map': 'CTY'})
        r = session.get(url, timeout=60)
        r.raise_for_status()
        county = COUNTIES_ALPHA[n - 1]
        for cid, (title, counts) in parse_counts(r.text).items():
            harvested.setdefault(cid, (title, {}))[1][county] = counts
        print(f'  cty {n:02d} ({county}): {len(parse_counts(r.text))} contests', flush=True)

    # sanity: contest ids must exist in the manifest (titles may differ in
    # whitespace only)
    unknown = set(harvested) - set(manifest)
    if unknown:
        raise SystemExit(f'contests on pages but not in manifest: {sorted(unknown)}')

    for cid, (title, by_county) in harvested.items():
        if re.sub(r'\s+', ' ', manifest[cid]['title']) != title:
            print(f'  note: title differs for {cid}: '
                  f'{manifest[cid]["title"]!r} vs page {title!r}')
        manifest[cid]['county_page_votes'] = by_county
    with open(manifest_path, 'w') as f:
        json.dump(manifest, f, indent=1)
    print(f'{tag}: county_page_votes added to {len(harvested)}/{len(manifest)} contests')


if __name__ == '__main__':
    import sys
    main(sys.argv[1], sys.argv[2])