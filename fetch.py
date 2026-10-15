"""Busca na API Trinks os dados necessários para o relatório de uma data de referência.
Uso: python3 fetch.py AAAA-MM-DD [estabelecimentoId ...]
A data de referência é o último dia fechado (ex.: relatório das 08:00 de 15/10 -> ref 2026-10-14).
Para o relatório das 14:00 do dia atual, use a data de hoje (dados parciais)."""
import sys, os, json, datetime as dt
from zoneinfo import ZoneInfo
import api

LOJAS = {184005: 'Lauro de Freitas', 158007: 'Vasco da Gama'}

def d(x): return x.strftime('%Y-%m-%dT00:00:00')
def e(x): return x.strftime('%Y-%m-%dT23:59:59')

def fetch(est, ref):
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'raw', str(est))
    os.makedirs(out, exist_ok=True)
    ini = ref.replace(day=1)
    prev_ini = (ini - dt.timedelta(days=1)).replace(day=1)
    hist_ini = (ini - dt.timedelta(days=95)).replace(day=1)
    try:
        prev_same = prev_ini.replace(day=ref.day)
    except ValueError:
        prev_same = ini - dt.timedelta(days=1)
    jobs = {
        'prof': ('/v1/profissionais', {}),
        'ag_mes': ('/v1/agendamentos', dict(dataInicio=d(ini), dataFim=e(ref))),
        'ag_hist': ('/v1/agendamentos', dict(dataInicio=d(hist_ini), dataFim=e(ini - dt.timedelta(days=1)))),
        'ag_fut': ('/v1/agendamentos', dict(dataInicio=d(ref + dt.timedelta(days=1)), dataFim=e(ref + dt.timedelta(days=45)))),
        'tr_mes': ('/v1/transacoes', dict(dataInicio=d(ini), dataFim=e(ref), incluirEstornos='false')),
        'tr_ant': ('/v1/transacoes', dict(dataInicio=d(prev_ini), dataFim=e(prev_same), incluirEstornos='false')),
        'clube': ('/v1/clube/assinaturas', {}),
        'cli_novos': ('/v1/clientes', dict(DataCadastroInicio=ini.isoformat(), DataCadastroFim=ref.isoformat())),
    }
    for name, (path, params) in jobs.items():
        data = api.get_all(path, est, pageSize=50, **params)
        json.dump(data, open(os.path.join(out, name + '.json'), 'w'), ensure_ascii=False)
    json.dump({'ref': ref.isoformat(), 'est': est, 'loja': LOJAS.get(est, str(est)),
               'gerado_em': dt.datetime.now(ZoneInfo('America/Sao_Paulo')).replace(tzinfo=None).isoformat(timespec='minutes')},
              open(os.path.join(out, 'meta.json'), 'w'))

if __name__ == '__main__':
    ref = dt.date.fromisoformat(sys.argv[1])
    ests = [int(x) for x in sys.argv[2:]] or list(LOJAS)
    for est in ests:
        try:
            fetch(est, ref); print('ok', est, LOJAS.get(est))
        except Exception as ex:
            print('ERRO', est, LOJAS.get(est), ex)
    print('chamadas', api.CALLS[0])
