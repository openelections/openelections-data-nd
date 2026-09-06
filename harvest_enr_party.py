# Harvest per-candidate party labels from the ENR site's county pages.
#
# Primary-election exports carry party in the contest title, but
# general-election contests ("President & Vice-President of the United
# States") have none — the parties appear only in the page HTML as
# candidate-row divs (<div class="Republican" >Republican<). This script
# GETs each county's results page, pairs each plain-candidate <h2> with
# the party div that follows it, and merges the (title, candidate) ->
# party map across counties into the election's manifest.json entries as
# 'candidate_parties'. Partly county-specific local contests keep their
# party-free (nonpartisan) entries.
#
# Usage: python3 harvest_enr_party.py <eid> <tag>
import json
import os
import re

from scrape_enr_exports import BASE, COUNTIES, OUT_DIR, election_url, session

PARTY_DIV = re.compile(r'<div class="([A-Za-z][A-Za-z -]*?)" ?>(?:\1)<')

# party labels as shown on the site -> repo convention (DEM/REP, blank
# for nonpartisan; Libertarian kept as LIB)
PARTY_FIXUPS = {
    'Democratic-NPL': 'DEM',
    'Democratic': 'DEM',
    'Republican': 'REP',
    'Libertarian': 'LIB',
}


def parse_parties(html):
    """Return {race title: {candidate: party}} for one county page."""
    out = {}
    for block in html.split('wrapper-inside wrapper-border')[1:]:
        title_m = re.search(r'<h1>(.*?)</h1>', block, re.S)
        if not title_m:
            continue
        title = re.sub(r'\s+', ' ', re.sub(r'<[^>]+>', ' ', title_m.group(1))).strip()
        candidates = [re.sub(r'\s+', ' ', c).strip()
                      for c in re.findall(r'<h2>([^<]+)</h2>', block)
                      if not re.fullmatch(r'Vote For \d+', c.strip())
                      and c.strip() != 'Total Votes']
        parties = [m.group(1) for m in PARTY_DIV.finditer(block)]
        if len(candidates) == len(parties):
            out.setdefault(title, {})
            out[title].update(zip(candidates, parties))
    return out


def main(eid, tag):
    election_dir = os.path.join(OUT_DIR, tag)
    manifest_path = os.path.join(election_dir, 'manifest.json')
    with open(manifest_path) as f:
        manifest = json.load(f)

    # title in the manifest -> title as harvested (cleaned the same way
    # the parser cleans contest names)
    from parse_enr_exports import parse_contest
    harvested = {}
    for n in COUNTIES:
        url = election_url(eid, 'ResultsSW.aspx', {'type': 'CTYALL', 'cty': f'{n:02d}', 'map': 'CTY'})
        r = session.get(url, timeout=60)
        r.raise_for_status()
        for title, cands in parse_parties(r.text).items():
            office, _, _ = parse_contest(title, '')
            harvested.setdefault(office, {}).update(
                {c: PARTY_FIXUPS.get(p, p) for c, p in cands.items()})
        print(f'  cty {n:02d}: {len(harvested)} contests seen', flush=True)

    added = 0
    for cid, info in manifest.items():
        office, _, _ = parse_contest(info['title'], '')
        parties = harvested.get(office)
        if parties:
            info['candidate_parties'] = parties
            added += 1
    with open(manifest_path, 'w') as f:
        json.dump(manifest, f, indent=1)
    print(f'{tag}: candidate_parties added to {added}/{len(manifest)} contests')


if __name__ == '__main__':
    import sys
    main(sys.argv[1], sys.argv[2])