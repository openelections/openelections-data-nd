# Scrape precinct-level election exports from North Dakota's retired
# "Election Night Reporting" site (results.sos.nd.gov), used for elections
# up to 2024. The 2026 election moved to resultsnd.sos.nd.gov, whose
# "download full data" export already lives in the sources repo.
#
# Each contest listed on a county's results page (ResultsSW.aspx?...cty=NN)
# has an "Export Precinct" button: an ASP.NET postback that returns an
# Excel workbook with one sheet per county and one row per precinct.
# The same contest appears on every county it is on the ballot of, so we
# enumerate all 53 county pages, dedupe contests, and download each once.
# A county-level "All Contests" CSV comes from ResultsExport.aspx.
#
# Usage: python3 scrape_enr_exports.py [--elections eid=tag,eid=tag ...]
import concurrent.futures
import os
import re
import sys
import time

import requests

BASE = 'https://results.sos.nd.gov/'
# ENR election IDs from the archived-elections list at
# https://www.sos.nd.gov/elections/election-results. The 2024 general is
# the site's default election and takes no eid parameter.
ELECTIONS = {
    'Ug32Itd2SDQ.': '20000613__primary',
    'LMRgBX6mD1A.': '20001107__general',
    'UMBawkGvl-U.': '20020611__primary',
    'uCduaMfbZTw.': '20021105__general',
    'NVuv10sW5FY.': '20040608__primary',
    'f4_9wSod8rs.': '20041102__general',
    'w3d71_B1QYk.': '20060613__primary',
    'YRm9d8aOerM.': '20061107__general',
    '3p15Rjra0is.': '20080603__primary',
    'PyBKp-DbstU.': '20081104__general',
    'UyRxWo3653c.': '20100608__primary',
    'SZ68n1F5X-M.': '20101102__general',
    'PJsJzM4LbMc.': '20120612__primary',
    'vp_Q9Zt7wF4.': '20121106__general',
    'Q5exdCNiAsM.': '20140610__primary',
    'XunCiWF0O8w.': '20141104__general',
    'YWxkP9oDgHk.': '20180612__primary',
    'bdXDzl1YpEQ.': '20181106__general',
    'StuhWbgeuSk.': '20201103__general',
    'ae0__NBsfOw.': '20200609__primary',
    'Next2sICxjI.': '20220614__primary',
    'vxUYQ0lrpP4.': '20221108__general',
    'OhimOTJtayc.': '20240611__primary',
    '': '20241105__general',
}
COUNTIES = range(1, 54)  # cty=01..53
OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                       '../openelections-sources-nd/nd_enr_exports')

session = requests.Session()
session.headers['User-Agent'] = 'Mozilla/5.0 (openelections-data-nd scraper)'


def election_url(eid, path, params):
    params = dict(params)
    if eid:
        params['eid'] = eid
    return BASE + path + '?' + '&'.join(f'{k}={v}' for k, v in params.items())


def form_fields(html):
    fields = {}
    for m in re.finditer(r'<input[^>]*name="([^"]+)"[^>]*>', html):
        name, tag = m.group(1), m.group(0)
        if not (name.startswith('__') or name.startswith('ctl00$hid')):
            continue
        val = re.search(r'value="([^"]*)"', tag)
        fields[name] = val.group(1) if val else ''
    return fields


def parse_races(html):
    """Return [(contest_id, title, precinct_postback_name, party)].

    Each race sits in its own `wrapper-inside wrapper-border` block with an
    <h1> title, a per-contest "Track this Contest" checkbox (chk-NNNN,
    unique to the contest including its party) and Precinct/County export
    buttons.
    """
    races = []
    for block in html.split('wrapper-inside wrapper-border')[1:]:
        title_m = re.search(r'<h1>(.*?)</h1>', block, re.S)
        chk_m = re.search(r'id="chk-(\d+)"', block)
        btn_m = re.search(
            r"__doPostBack\('(ctl00\$MainContent\$rptRace[^']+)',''\)"
            r'" name="[^"]*" type="button" class="export-button" '
            r'value="Precinct" raceid="\d+" party="([^"]*)"', block)
        if not (title_m and chk_m and btn_m):
            continue
        title = re.sub(r'\s+', ' ', re.sub(r'<[^>]+>', ' ', title_m.group(1))).strip()
        races.append((chk_m.group(1), title, btn_m.group(1), btn_m.group(2)))
    return races


def enumerate_races(eid):
    """GET each county page; return (page_info, unique races)."""
    pages = {}
    races = {}
    for n in COUNTIES:
        url = election_url(eid, 'ResultsSW.aspx', {'type': 'CTYALL', 'cty': f'{n:02d}', 'map': 'CTY'})
        r = session.get(url, timeout=60)
        r.raise_for_status()
        pages[n] = {'url': url, 'fields': form_fields(r.text)}
        for race in parse_races(r.text):
            races.setdefault(race[0], {'race': race, 'cty': n})
        print(f'  cty {n:02d}: {len(races)} races on page, {len(races)} unique so far', flush=True)
    return pages, races


def download_precinct_export(page, target, out_path):
    fields = dict(page['fields'])
    fields['__EVENTTARGET'] = target
    fields['__EVENTARGUMENT'] = ''
    for attempt in range(3):
        try:
            r = session.post(page['url'], data=fields, timeout=120)
            if 'spreadsheet' in r.headers.get('Content-Type', ''):
                with open(out_path, 'wb') as f:
                    f.write(r.content)
                return len(r.content)
            # error page: some races have no export on the old platform
            if 'Error.aspx' in r.text[:2000]:
                return None
            time.sleep(2 * (attempt + 1))
        except requests.RequestException:
            time.sleep(2 * (attempt + 1))
    return None  # give up on this race; retried after the main pass


def main():
    only = sys.argv[sys.argv.index('--elections') + 1].split(',') if '--elections' in sys.argv else ELECTIONS
    for eid in only:
        tag = ELECTIONS[eid]
        out = os.path.join(OUT_DIR, tag)
        os.makedirs(os.path.join(out, 'precinct'), exist_ok=True)
        print(f'{tag} (eid={eid or "default"}): enumerating county pages', flush=True)
        pages, races = enumerate_races(eid)
        # manifest: contest_id -> title/party, consumed by the parser (the
        # workbooks themselves don't carry the party)
        import json
        with open(os.path.join(out, 'manifest.json'), 'w') as f:
            json.dump({cid: {'title': info['race'][1],
                             'party': info['race'][3],
                             'precinct_postback': info['race'][2]}
                       for cid, info in races.items()}, f, indent=1)
        print(f'{tag}: {len(races)} unique contests; downloading precinct exports', flush=True)

        def job(contest_id):
            info = races[contest_id]
            race = info['race']
            path = os.path.join(out, 'precinct', f'{contest_id}.xlsx')
            if os.path.exists(path):
                return contest_id, -1  # already downloaded on a previous run
            size = download_precinct_export(pages[info['cty']], race[2], path)
            return contest_id, size

        done = 0
        failures = []
        with concurrent.futures.ThreadPoolExecutor(max_workers=6) as ex:
            for contest_id, size in ex.map(job, list(races)):
                done += 1
                if size is None:
                    failures.append(contest_id)
                if done % 25 == 0:
                    print(f'  {done}/{len(races)} done', flush=True)
        # retry the failures once, sequentially
        retried = []
        for contest_id in list(failures):
            info = races[contest_id]
            path = os.path.join(out, 'precinct', f'{contest_id}.xlsx')
            size = download_precinct_export(pages[info['cty']],
                                            info['race'][2], path)
            if size is None:
                retried.append(contest_id)
                failures.remove(contest_id)
        if retried:
            print(f'{tag}: {len(retried)} exports unavailable after retry: '
                  f'{retried[:10]}{"..." if len(retried) > 10 else ""}', flush=True)

        # County-level all-contests CSV from ResultsExport.aspx (used only
        # as a cross-check; tolerate platforms that don't offer it)
        try:
            url = election_url(eid, 'ResultsExport.aspx', {})
            page = session.get(url, timeout=60)
            fields = form_fields(page.text)
            fields['ctl00$MainContent$rblTypes'] = '1'
            fields['ctl00$MainContent$ctl01'] = 'All Contests'
            r = session.post(url, data=fields, timeout=300)
            assert 'text' in r.headers.get('Content-Type', ''), r.headers.get('Content-Type')
            with open(os.path.join(out, 'all_contests_county.csv'), 'wb') as f:
                f.write(r.content)
            print(f'{tag}: done ({len(races)} precinct exports + county CSV)', flush=True)
        except Exception as e:
            print(f'{tag}: done ({len(races)} precinct exports, county CSV unavailable: {e})',
                  flush=True)


if __name__ == '__main__':
    main()