import json, time, os, sys, urllib.request, urllib.parse
_kf=os.path.join(os.path.dirname(os.path.abspath(__file__)),'.key')
KEY=(os.environ.get('TRINKS_KEY') or (open(_kf).read() if os.path.exists(_kf) else '')).strip()
BASE='https://api.trinks.com'
CALLS=[0]
def get(path, est=None, **params):
    q=urllib.parse.urlencode({k:v for k,v in params.items() if v is not None})
    url=f"{BASE}{path}"+(f"?{q}" if q else "")
    h={'X-Api-Key':KEY,'Accept':'application/json'}
    if est: h['estabelecimentoId']=str(est)
    for attempt in range(5):
        req=urllib.request.Request(url,headers=h)
        try:
            CALLS[0]+=1; time.sleep(1.2)
            with urllib.request.urlopen(req,timeout=60) as r:
                return json.loads(r.read() or b'null')
        except urllib.error.HTTPError as e:
            body=e.read().decode(errors='ignore')
            if e.code==429:
                wait=int(e.headers.get('Retry-After') or 30); print('429 wait',wait,body[:200],file=sys.stderr); time.sleep(wait); continue
            raise RuntimeError(f"HTTP {e.code} {url}: {body[:500]}")
    raise RuntimeError('too many retries')
def get_all(path, est=None, pageSize=100, **params):
    out=[]; page=1
    while True:
        r=get(path, est, page=page, pageSize=pageSize, **params)
        if isinstance(r,list): return r
        out+=r.get('data') or []
        if page>=(r.get('totalPages') or 1): return out
        page+=1
