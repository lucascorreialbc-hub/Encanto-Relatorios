"""Relatório automático Encanto / Trinks.
Uso: TRINKS_KEY=... python3 run.py manha|tarde [AAAA-MM-DD]   (data = "hoje"; padrão: hoje em São Paulo)
manha: fecha o dia anterior e planeja o dia (segunda-feira inclui a visão semanal)
tarde: parcial do dia atual
Saída: out/<arquivo>.xlsx e out/email.json (assunto, html, texto, caminho da planilha, avisos)."""
import sys, os, glob, json, subprocess, datetime as dt
from zoneinfo import ZoneInfo
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
import fetch, build, mail

def recalc(path):
    cands = glob.glob('/root/.claude/skills/**/xlsx/scripts/recalc.py', recursive=True) + \
            glob.glob('/tmp/**/xlsx/scripts/recalc.py', recursive=True)
    if not cands: return {'error': 'recalc.py não encontrado'}
    r = subprocess.run([sys.executable, cands[0], path, '90'], capture_output=True, text=True)
    try: return json.loads(r.stdout[r.stdout.index('{'):])
    except Exception: return {'error': (r.stdout + r.stderr)[-500:]}

def main():
    modo = sys.argv[1] if len(sys.argv) > 1 else 'manha'
    hoje = dt.date.fromisoformat(sys.argv[2]) if len(sys.argv) > 2 else dt.datetime.now(ZoneInfo('America/Sao_Paulo')).date()
    ref = hoje - dt.timedelta(days=1) if modo == 'manha' else hoje
    semanal = modo == 'manha' and hoje.weekday() == 0
    avisos, stores = [], []
    os.makedirs(os.path.join(HERE, 'out'), exist_ok=True)
    for est, nome in fetch.LOJAS.items():
        try:
            fetch.fetch(est, ref)
            d = build.load(os.path.join(HERE, 'raw', str(est)))
            st = build.prepare(d); st['insights'] = build.insights(st); stores.append(st)
        except Exception as ex:
            msg = str(ex)
            if 'não liberado' in msg or 'n\\u00E3o liberado' in msg or '403' in msg:
                avisos.append(f'{nome}: a Trinks ainda não liberou a API desta unidade.')
            else:
                avisos.append(f'{nome}: erro ao buscar dados ({msg[:150]}).')
    if not stores:
        json.dump({'erro': True, 'avisos': avisos}, open(os.path.join(HERE, 'out', 'email.json'), 'w'), ensure_ascii=False)
        print('ERRO', avisos); return
    os.makedirs(os.path.join(HERE, 'out'), exist_ok=True)
    tag = {'manha': 'Manha', 'tarde': 'Tarde'}[modo]
    xlsx = os.path.join(HERE, 'out', f'Encanto_{tag}_{hoje:%Y-%m-%d}.xlsx')
    build.build(xlsx, stores, modo=modo, semanal=semanal)
    rc = recalc(xlsx)
    if rc.get('status') != 'success': avisos.append(f'Recalculo da planilha: {rc}')
    em = mail.compose(xlsx, stores, modo, hoje, semanal, avisos)
    em['xlsx'] = xlsx; em['avisos'] = avisos; em['chamadas_api'] = fetch.api.CALLS[0]
    json.dump(em, open(os.path.join(HERE, 'out', 'email.json'), 'w'), ensure_ascii=False)
    agora = dt.datetime.now(ZoneInfo('America/Sao_Paulo')).strftime('%d/%m/%Y %H:%M')
    open(os.path.join(HERE, 'out', 'painel.html'), 'w').write(mail.page(em, agora))
    print('ok', xlsx, '| chamadas API:', fetch.api.CALLS[0], '| avisos:', avisos)

if __name__ == '__main__':
    main()
