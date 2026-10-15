"""Monta o dashboard Excel a partir dos dados baixados por fetch.py.
Uso: python3 build.py saida.xlsx raw/184005 [raw/158007 ...]"""
import sys, os, json, datetime as dt, collections as C
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import DataBarRule
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L

# ---------------------------------------------------------------- configuração
FERIADOS = {dt.date(2026, 10, 12), dt.date(2026, 11, 2), dt.date(2026, 11, 20), dt.date(2026, 12, 25)}
META_MES = {'2026-10': 50000}                 # meta por unidade (informada pelo Lucas)
LINHAS = [('Alisamentos', 30000), ('Assinatura Encanto', 8000), ('Tratamentos e escovas', 5000),
          ('Manicure', 4000), ('Home care', 3000)]
DIAS_SUMIDA = 90                  # cliente "sumida": 90+ dias sem atendimento finalizado
ALISAMENTOS = {'MAJESTADE', 'PROGRESSIVA PRINCESA', 'PRINCESA'}
SIGLA = {184005: 'LF', 158007: 'VG'}
# equipe por loja: id (agenda) -> (nome de exibição, função, meta principal, valor da meta, tipo de medida)
EQUIPE = {
    184005: {
        663619: ('Lívia', 'Recepção', 'Clientes na loja', 160, 'visitas'),
        715590: ('Daiane Oliveira', 'Closer', 'Resgates da loja', 40, 'resgates'),
        917264: ('Julia Costa', 'Especialista', 'Atendimentos no mês', 55, 'atend'),
        923873: ('Rosana Rodrigues', 'Especialista', 'Atendimentos no mês', 55, 'atend'),
        944846: ('Laiane Lima', 'Manicure', 'Faturamento de serviços', 4000, 'fat'),
        621335: ('Tamiles', 'Gerência', None, None, None),
        652782: ('Lucas', 'Gerência', None, None, None),
        628345: ('Mirlene', 'Recepção', None, None, None),
    },
}
DIAS = ['seg', 'ter', 'qua', 'qui', 'sex', 'sáb', 'dom']

# ---------------------------------------------------------------- estilo
CORAL, CORAL_L, INK, INK2, MUTED, LINE, BG, WHITE, GREEN, RED = (
    'D9644A', 'FBE9E4', '2A2A33', '55556A', '8A8A99', 'E4E2E8', 'F7F6F8', 'FFFFFF', '2E7D5B', 'B23A3A')
F = lambda **k: Font(name='Arial', **k)
fill = lambda c: PatternFill('solid', fgColor=c)
thin = Side(style='thin', color=LINE)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
BOT = Border(bottom=thin)
BRL = '"R$" #,##0;-"R$" #,##0;"–"'
BRL2 = '"R$" #,##0.00;-"R$" #,##0.00;"–"'
INT = '#,##0;-#,##0;"–"'
PCT = '0%;-0%;"–"'
DT = 'dd/mm'

def hdr(ws, row, cols, start=1):
    for i, h in enumerate(cols):
        c = ws.cell(row, start + i, h)
        c.font = F(bold=True, color=WHITE, size=10); c.fill = fill(INK)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[row].height = 30

def widths(ws, ws_w):
    for i, w in enumerate(ws_w, 1): ws.column_dimensions[L(i)].width = w

def title(ws, text, sub, ncols):
    ws.sheet_view.showGridLines = False
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    a = ws.cell(1, 1, text); a.font = F(bold=True, size=16, color=WHITE); a.fill = fill(INK)
    a.alignment = Alignment(vertical='center', indent=1)
    b = ws.cell(2, 1, sub); b.font = F(size=10, color=WHITE); b.fill = fill(INK)
    b.alignment = Alignment(vertical='center', indent=1)
    for c in range(1, ncols + 1):
        ws.cell(1, c).fill = fill(INK); ws.cell(2, c).fill = fill(INK)
    ws.row_dimensions[1].height = 30; ws.row_dimensions[2].height = 20

def table(ws, r0, header, rows, fmts, start=1, zebra=True):
    hdr(ws, r0, header, start)
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            c = ws.cell(r0 + 1 + i, start + j, v)
            c.font = F(size=10, color=INK); c.border = BOT
            if fmts[j]: c.number_format = fmts[j]
            if zebra and i % 2: c.fill = fill(BG)
    return r0 + len(rows)

# ---------------------------------------------------------------- dados
def load(path):
    j = lambda n: json.load(open(os.path.join(path, n + '.json')))
    return {n: j(n) for n in ['meta', 'prof', 'ag_mes', 'ag_hist', 'ag_fut', 'tr_mes', 'tr_ant', 'clube', 'cli_novos']}

def pdate(s): return dt.date.fromisoformat(s[:10])

def linha_servico(s):
    if s['nome'].strip().upper() in ALISAMENTOS: return 'Alisamentos'
    cat = (s.get('categoria') or '').lower()
    if any(k in cat for k in ('mão', 'mao', 'pé', 'manicure', 'pedicure', 'unha')): return 'Manicure'
    return 'Tratamentos e escovas'

def workdays(y, m):
    d = dt.date(y, m, 1); out = []
    while d.month == m:
        out.append(d); d += dt.timedelta(days=1)
    return out

def is_util(d): return d.weekday() < 6 and d not in FERIADOS

def prepare(D):
    est = D['meta']['est']; ref = dt.date.fromisoformat(D['meta']['ref']); sig = SIGLA.get(est, str(est))
    cfg = EQUIPE.get(est, {})
    byid, byidp = {}, {}
    for p in D['prof']:
        nome = cfg.get(p['id'], (p.get('apelido') or p['nome'].strip(),))[0]
        byid[p['id']] = nome; byidp[p['idProfissional']] = nome
    nome_ag = lambda a: byid.get((a.get('profissional') or {}).get('id'), 'Sem profissional') if a.get('profissional') else 'Sem profissional'

    # ----- itens de receita (transações)
    itens, pags = [], []
    for t in sorted(D['tr_mes'], key=lambda t: t['dataHora']):
        dia = pdate(t['dataHora']); cli = (t.get('cliente') or {}).get('nome', '').strip()
        lines = []
        for s in t['servicos']:
            lines.append(['Serviço', s['nome'].strip(), linha_servico(s), byidp.get(s.get('idProfissionalQueRealizouServico'), 'Sem profissional'), 1, s['preco']])
        for p in t['produtos']:
            lines.append(['Produto', p['nome'].strip(), 'Home care', byidp.get(p.get('IdProfissionalQueRealizouAVenda'), 'Sem profissional'), p['quantidade'], p['quantidade'] * p['valorUnitario']])
        for p in t['pacotes']:
            lines.append(['Pacote', p['nome'].strip(), 'Tratamentos e escovas', byidp.get(p.get('IdProfissionalQueRealizouAVenda'), 'Sem profissional'), p['quantidade'], p['quantidade'] * p['valorUnitario']])
        bruto = sum(l[5] for l in lines)
        if not lines:   # cobrança recorrente do clube
            lines = [['Assinatura', 'Mensalidade do clube', 'Assinatura Encanto', 'Clube (cobrança automática)', 1, t['totalPagar']]]
            bruto = t['totalPagar']
        desc = t['totalPagar'] - bruto   # negativo quando há desconto
        seen500 = set()
        rats = [round(desc * l[5] / bruto, 2) if bruto else 0 for l in lines]
        rats[-1] = round(desc - sum(rats[:-1]), 2)
        for k, l in enumerate(lines):
            rat = rats[k]
            flag = 0
            if t['totalPagar'] >= 500 and l[0] == 'Serviço' and l[3] not in seen500:
                flag = 1; seen500.add(l[3])
            itens.append([sig, dia, t['id'], cli] + l[:4] + [l[4], l[5], rat, None, t['totalPagar'], flag])
        for fp in t['formasPagamentos']:
            pags.append([sig, dia, t['id'], cli, fp['nome'], fp['valor']])
    # acerto de centavos do rateio para que a soma líquida bata com o total pago
    for it in itens: it[11] = round(it[9] + it[10], 2)

    # ----- agenda
    novos = {c['id'] for c in D['cli_novos']}
    ev = C.defaultdict(list)   # cliente -> [(data, status)]
    for a in D['ag_hist'] + D['ag_mes'] + D['ag_fut']:
        if a.get('cliente'): ev[a['cliente']['id']].append((dt.datetime.fromisoformat(a['dataHoraInicio']), a['status']['nome']))
    for k in ev: ev[k].sort()
    FIN, FALTA, CANC = 'Finalizado', 'Cliente não compareceu', 'Cancelado'
    agenda, seen_visit, seen_new = [], set(), set()
    for a in sorted(D['ag_mes'], key=lambda a: a['dataHoraInicio']):
        when = dt.datetime.fromisoformat(a['dataHoraInicio']); dia = when.date()
        cid = (a.get('cliente') or {}).get('id'); st = a['status']['nome']
        visita = nova = resg = ret = 0
        if st == FIN and cid and (cid, dia) not in seen_visit:
            seen_visit.add((cid, dia)); visita = 1
            if cid in novos and cid not in seen_new: nova = 1; seen_new.add(cid)
            prior = [(w, s) for w, s in ev[cid] if w.date() < dia]
            last_fin = max([w for w, s in prior if s == FIN], default=None)
            after = [(w, s) for w, s in prior if last_fin is None or w > last_fin]
            falhou = any(s == FALTA or (s == CANC and (dia - w.date()).days >= 7) for w, s in after)
            if cid not in novos and (last_fin is None or (dia - last_fin.date()).days >= DIAS_SUMIDA or falhou): resg = 1
            ret = int(any(w.date() > dia and s != CANC for w, s in ev[cid]))
        agenda.append([sig, dia, when.strftime('%H:%M'), (a.get('cliente') or {}).get('nome', '').strip(),
                       a['servico']['nome'].strip(), nome_ag(a), st, a.get('valor') or 0, visita, nova, resg, ret])
    futuro = [[sig, dt.datetime.fromisoformat(a['dataHoraInicio']).date(), a['dataHoraInicio'][11:16],
               (a.get('cliente') or {}).get('nome', '').strip(), a['servico']['nome'].strip(), nome_ag(a),
               a['status']['nome'], a.get('valor') or 0]
              for a in sorted(D['ag_fut'], key=lambda a: a['dataHoraInicio']) if a['status']['nome'] != CANC]

    # ----- desistências do mês e sumidas
    desist = []
    for a in sorted(D['ag_mes'], key=lambda a: a['dataHoraInicio']):
        st = a['status']['nome']
        if st not in (CANC, FALTA) or not a.get('cliente'): continue
        cid = a['cliente']['id']; w0 = dt.datetime.fromisoformat(a['dataHoraInicio'])
        voltou = [w for w, s in ev[cid] if w > w0 and w.date() <= ref and s == FIN]
        marc = [w for w, s in ev[cid] if w.date() > ref and s != CANC]
        sit = f'Voltou em {voltou[0]:%d/%m}' if voltou else (f'Remarcada para {marc[0]:%d/%m}' if marc else 'Pendente – resgatar')
        desist.append([w0.date(), a['cliente']['nome'].strip(), a['servico']['nome'].strip(), nome_ag(a), st, sit])
    sumidas = []
    nomecli = {}
    ultimo = {}
    for a in D['ag_hist'] + D['ag_mes']:
        if a.get('cliente'): nomecli[a['cliente']['id']] = a['cliente']['nome'].strip()
        if a['status']['nome'] == FIN and a.get('cliente'):
            w = dt.datetime.fromisoformat(a['dataHoraInicio'])
            cid = a['cliente']['id']
            if cid not in ultimo or w > ultimo[cid][0]: ultimo[cid] = (w, a['servico']['nome'].strip(), nome_ag(a))
    nvis = C.Counter()
    for a in D['ag_hist'] + D['ag_mes']:
        if a['status']['nome'] == FIN and a.get('cliente'): nvis[(a['cliente']['id'], a['dataHoraInicio'][:10])] = 1
    nv = C.Counter(c for c, _ in nvis)
    for cid, (w, serv, prof) in ultimo.items():
        dias = (ref - w.date()).days
        if dias >= 60 and not any(x.date() > ref and s != CANC for x, s in ev[cid]):
            sumidas.append([nomecli[cid], w.date(), dias, serv, prof, nv[cid]])
    sumidas.sort(key=lambda r: r[2])

    # ----- clube
    ini = ref.replace(day=1)
    clube = []
    for s in D['clube']:
        da = pdate(s['dataAssinatura']); dc = pdate(s['dataCancelamento']) if s.get('dataCancelamento') else None
        vend = byid.get(s.get('profissionalDeVendaId')) or ('Site / online' if s.get('canalDeVenda') == 'Site' else 'Não informado')
        if ini <= da <= ref:
            clube.append([sig, 'Nova', da, s['clienteNome'].strip(), s['nome'].strip(), s['status'], s['valorCobranca'], vend])
        if dc and ini <= dc <= ref:
            clube.append([sig, 'Cancelada', dc, s['clienteNome'].strip(), s['nome'].strip(), s['status'], s['valorCobranca'], vend])
    ativas = sum(1 for s in D['clube'] if s['status'] == 'Ativa')
    ant = sum(t['totalPagar'] for t in D['tr_ant'])
    return dict(est=est, sig=sig, loja=D['meta']['loja'], ref=ref, itens=itens, pags=pags, agenda=agenda,
                futuro=futuro, desist=desist, sumidas=sumidas, clube=clube, ativas=ativas, ant=ant,
                novos=len(D['cli_novos']), cfg=cfg, gerado=D['meta']['gerado_em'],
                pessoas=sorted({i[7] for i in itens} | {a[5] for a in agenda}))

# ---------------------------------------------------------------- planilha
def build(out, stores, modo='manha', semanal=False):
    wb = Workbook(); wb.remove(wb.active)
    ref = stores[0]['ref']; mes = f'{ref:%Y-%m}'
    dias_mes = workdays(ref.year, ref.month)
    prox = ref + dt.timedelta(days=1)
    while not is_util(prox): prox += dt.timedelta(days=1)
    refu = ref                                   # último dia útil com dados (tiles do "dia")
    while not is_util(refu) and refu.day > 1: refu -= dt.timedelta(days=1)

    # ---- Premissas
    P = wb.create_sheet('Premissas')
    title(P, 'Premissas e metas', 'Células em azul são entradas: altere aqui e todo o painel recalcula.', 2 + len(stores))
    widths(P, [44] + [18] * len(stores) + [60])
    rows = [('Loja', [s['loja'] for s in stores], None, None),
            ('Sigla', [s['sig'] for s in stores], None, None),
            ('Data de referência (último dia com dados)', [ref] * len(stores), 'dd/mm/yyyy', 'Último dia incluído no relatório.'),
            ('Próximo dia útil (o "hoje" do plano)', [prox] * len(stores), 'dd/mm/yyyy', 'Dia para o qual o painel calcula a necessidade de venda.'),
            ('Último dia útil fechado', [refu] * len(stores), 'dd/mm/yyyy', 'Dia mostrado no bloco do dia (no relatório das 14h, é o próprio dia, parcial).'),
            ('Meta do mês (R$)', [META_MES.get(mes, 50000)] * len(stores), BRL, 'Meta de outubro informada pelo Lucas: R$ 50 mil por unidade.'),
            ('Gatilho do bônus da equipe (R$)', [45000] * len(stores), BRL, 'Regra da casa: bônus liberado a partir de R$ 45 mil.'),
            ('Dias úteis no mês', None, INT, 'Segunda a sábado, menos feriados (aba Diário).'),
            ('Dias úteis decorridos', None, INT, None),
            ('Dias úteis restantes', None, INT, None),
            ('Faturamento mês anterior até o mesmo dia (R$)', [s['ant'] for s in stores], BRL, 'Soma do total pago nas transações da Trinks, mesmo período do mês anterior.'),
            ('Clientes novas cadastradas no mês', [s['novos'] for s in stores], INT, 'Trinks: clientes com data de cadastro no mês.'),
            ('Assinaturas ativas hoje', [s['ativas'] for s in stores], INT, 'Trinks: Clube de Assinaturas com status Ativa.'),
            ('Meta clientes na loja (recepção)', [160] * len(stores), INT, 'Plano de metas de outubro.'),
            ('Meta clientes com retorno marcado', [60] * len(stores), INT, 'Plano de metas de outubro.'),
            ('Meta resgates (closer da unidade)', [40] * len(stores), INT, 'Plano de metas de outubro: 40 resgates por closer, uma em cada unidade.')]
    hdr(P, 4, ['Item'] + [s['sig'] for s in stores] + ['Fonte / observação'])
    PR = {}
    for i, (lab, vals, fmt, note) in enumerate(rows):
        r = 5 + i; PR[lab] = r
        P.cell(r, 1, lab).font = F(size=10, color=INK)
        for j in range(len(stores)):
            c = P.cell(r, 2 + j)
            if vals is not None:
                c.value = vals[j]; c.font = F(size=10, color='0000FF')
            if fmt: c.number_format = fmt
        if note: P.cell(r, 2 + len(stores), note).font = F(size=9, color=MUTED)
        for cc in range(1, 3 + len(stores)): P.cell(r, cc).border = BOT
    r = 5 + len(rows) + 1
    P.cell(r, 1, 'Metas por linha de receita (R$ no mês)').font = F(bold=True, size=11, color=INK)
    hdr(P, r + 1, ['Linha de receita', 'Meta'])
    LR0 = r + 2
    for i, (n, v) in enumerate(LINHAS):
        P.cell(LR0 + i, 1, n).font = F(size=10)
        c = P.cell(LR0 + i, 2, v); c.font = F(size=10, color='0000FF'); c.number_format = BRL
    P.cell(LR0 + len(LINHAS), 1, 'Fonte: plano de metas de outubro (premissa de ticket médio de alisamento R$ 400).').font = F(size=9, color=MUTED, italic=True)
    r = LR0 + len(LINHAS) + 2
    P.cell(r, 1, 'Feriados considerados').font = F(bold=True, size=11, color=INK)
    for i, fd in enumerate(sorted(FERIADOS)):
        c = P.cell(r + 1 + i, 1, fd); c.number_format = 'dd/mm/yyyy'; c.font = F(size=10, color=INK)
    P.cell(r, 2, '(usados para montar o calendário de dias úteis; mudanças valem a partir do próximo relatório)').font = F(size=9, color=MUTED, italic=True)
    pref = lambda lab, j: f"Premissas!${L(2+j)}${PR[lab]}"

    # ---- bases de dados
    B = wb.create_sheet('Base_Receita')
    itens = [i for s in stores for i in s['itens']]
    bh = ['Loja', 'Data', 'Transação', 'Cliente', 'Tipo', 'Item', 'Linha de receita', 'Profissional', 'Qtd', 'Valor bruto', 'Desconto rateado', 'Valor líquido', 'Total da conta', 'Conta ≥ R$500 (1 por prof.)']
    nB = table(B, 1, bh, itens, [None, DT, None, None, None, None, None, None, INT, BRL2, BRL2, BRL2, BRL2, INT])
    widths(B, [6, 8, 11, 28, 10, 26, 20, 24, 5, 11, 11, 11, 11, 12]); B.freeze_panes = 'A2'; B.auto_filter.ref = f'A1:N{nB}'
    B.cell(1, 11).comment = Comment('Desconto da conta (ex.: benefício do clube) rateado entre os itens pelo valor bruto.', 'Claude')
    Bc = lambda col: f'Base_Receita!${col}$2:${col}${max(nB, 2)}'

    PG = wb.create_sheet('Base_Pagamentos')
    pags = [p for s in stores for p in s['pags']]
    nP = table(PG, 1, ['Loja', 'Data', 'Transação', 'Cliente', 'Forma de pagamento', 'Valor'], pags, [None, DT, None, None, None, BRL2])
    widths(PG, [6, 8, 11, 30, 30, 12]); PG.freeze_panes = 'A2'; PG.auto_filter.ref = f'A1:F{nP}'
    Pc = lambda col: f'Base_Pagamentos!${col}$2:${col}${max(nP, 2)}'

    AG = wb.create_sheet('Base_Agenda')
    ag = [a for s in stores for a in s['agenda']]
    nA = table(AG, 1, ['Loja', 'Data', 'Hora', 'Cliente', 'Serviço', 'Profissional', 'Status', 'Valor', 'Visita', 'Nova cliente', 'Resgate', 'Retorno marcado'], ag,
               [None, DT, None, None, None, None, None, BRL2, INT, INT, INT, INT])
    widths(AG, [6, 8, 7, 30, 24, 24, 22, 10, 8, 9, 9, 10]); AG.freeze_panes = 'A2'; AG.auto_filter.ref = f'A1:L{nA}'
    AG.cell(1, 9).comment = Comment('1 = primeira passagem finalizada da cliente naquele dia (conta como cliente na loja).', 'Claude')
    AG.cell(1, 11).comment = Comment('1 = cliente que voltou e foi atendida após 60+ dias sem vir, sem visita desde o início do histórico, ou após faltar / cancelar há 7+ dias. Clientes novas não contam.', 'Claude')
    AG.cell(1, 12).comment = Comment('1 = no dia do relatório a cliente já tem outro horário marcado depois desta visita.', 'Claude')
    Ac = lambda col: f'Base_Agenda!${col}$2:${col}${max(nA, 2)}'

    CL = wb.create_sheet('Base_Clube')
    cl = [c for s in stores for c in s['clube']]
    nC = table(CL, 1, ['Loja', 'Movimento', 'Data', 'Cliente', 'Plano', 'Status atual', 'Valor/mês', 'Vendido por'], cl, [None, None, DT, None, None, None, BRL, None])
    widths(CL, [6, 11, 8, 32, 22, 22, 10, 22]); CL.freeze_panes = 'A2'
    Cc = lambda col: f'Base_Clube!${col}$2:${col}${max(nC, 2)}'

    # ---- Diário (uma linha por dia útil do mês, colunas por loja)
    DI = wb.create_sheet('Diário')
    uteis = [d for d in dias_mes if is_util(d)]
    title(DI, 'Faturamento dia a dia (dias úteis)', 'Meta do dia = meta do mês ÷ dias úteis. Valores de domingo/feriado (ex.: cobranças do clube) entram no dia útil seguinte.', 2 + 5 * len(stores))
    head = ['Data', 'Dia']
    for s in stores: head += [f'{s["sig"]} · Realizado', f'{s["sig"]} · Meta do dia', f'{s["sig"]} · Acumulado', f'{s["sig"]} · Meta acumulada', f'{s["sig"]} · Atendimentos']
    hdr(DI, 4, head)
    D0 = 5; D1 = D0 + len(uteis) - 1
    DREF = D0 + max(0, sum(1 for d in uteis if d <= ref) - 1)   # última linha com realizado
    for i, d in enumerate(uteis):
        r = D0 + i
        DI.cell(r, 1, d).number_format = 'dd/mm'; DI.cell(r, 2, f'{d:%d} {DIAS[d.weekday()]}')
        for j, s in enumerate(stores):
            c0 = 3 + 5 * j; rc = pref('Data de referência (último dia com dados)', j); sg = f'"{s["sig"]}"'
            low = f'{Bc("B")},">"&A{r-1}' if i else f'{Bc("B")},">="&DATE(YEAR(A{r}),MONTH(A{r}),1)'
            DI.cell(r, c0, f'=IF(A{r}>{rc},"",SUMIFS({Bc("L")},{Bc("A")},{sg},{low},{Bc("B")},"<="&A{r}))')
            DI.cell(r, c0 + 1, f'={pref("Meta do mês (R$)", j)}/{pref("Dias úteis no mês", j)}')
            DI.cell(r, c0 + 2, f'=IF(A{r}>{rc},"",SUMIFS({Bc("L")},{Bc("A")},{sg},{Bc("B")},"<="&A{r}))')
            DI.cell(r, c0 + 3, f'=SUM(${L(c0+1)}${D0}:{L(c0+1)}{r})')
            DI.cell(r, c0 + 4, f'=IF(A{r}>{rc},"",COUNTIFS({Ac("A")},{sg},{Ac("B")},A{r},{Ac("G")},"Finalizado"))')
            for k, f in enumerate([BRL, BRL, BRL, BRL, INT]): DI.cell(r, c0 + k).number_format = f
        for c in range(1, len(head) + 1):
            DI.cell(r, c).font = F(size=10, color=INK, bold=(r == DREF)); DI.cell(r, c).border = BOT
            if r == DREF: DI.cell(r, c).fill = fill(CORAL_L)
    widths(DI, [8, 8] + [13] * 5 * len(stores)); DI.freeze_panes = 'C5'
    for j in range(len(stores)):
        P.cell(PR['Dias úteis no mês'], 2 + j, f'=COUNT(Diário!$A${D0}:$A${D1})')
        P.cell(PR['Dias úteis decorridos'], 2 + j, f'=COUNTIF(Diário!$A${D0}:$A${D1},"<="&{pref("Data de referência (último dia com dados)", j)})')
        P.cell(PR['Dias úteis restantes'], 2 + j, f'={pref("Dias úteis no mês", j)}-{pref("Dias úteis decorridos", j)}')

    # ---- por loja: Painel, Equipe, Ação
    for j, s in enumerate(stores):
        sg = f'"{s["sig"]}"'; pr = lambda lab: pref(lab, j)
        REF = pr('Data de referência (último dia com dados)'); PROX = pr('Próximo dia útil (o "hoje" do plano)')
        META = pr('Meta do mês (R$)')
        W = wb.create_sheet(f'Painel {s["sig"]}', j * 3)
        title(W, f'Encanto Progressivas · {s["loja"]}', f'Painel de {ref:%m/%Y} · dados da Trinks até {ref:%d/%m/%Y} · gerado em {s["gerado"][:16].replace("T", " ")}', 13)
        widths(W, [2] + [12.5] * 12 + [2])
        FAT = f'SUMIFS({Bc("L")},{Bc("A")},{sg})'
        def tile(r, c, label, formula, fmt, sub=None, accent=False, sub_fmt=None):
            for rr in range(r, r + 3):
                for cc in (c, c + 1):
                    x = W.cell(rr, cc); x.fill = fill(CORAL_L if accent else BG)
            W.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + 1)
            W.merge_cells(start_row=r + 1, start_column=c, end_row=r + 1, end_column=c + 1)
            W.merge_cells(start_row=r + 2, start_column=c, end_row=r + 2, end_column=c + 1)
            a = W.cell(r, c, label); a.font = F(size=9, color=INK2, bold=True); a.alignment = Alignment(indent=1, vertical='bottom')
            b = W.cell(r + 1, c, formula); b.font = F(size=18, bold=True, color=CORAL if accent else INK); b.number_format = fmt
            b.alignment = Alignment(indent=1, vertical='center', horizontal='left')
            if sub:
                if isinstance(sub, str) and sub.startswith('=') and '"R$ #,##0"' in sub:
                    sub = '=SUBSTITUTE(' + sub[1:] + ',",",".")'   # separador de milhar brasileiro em qualquer idioma
                x = W.cell(r + 2, c, sub); x.font = F(size=9, color=MUTED); x.alignment = Alignment(indent=1, vertical='top')
                if sub_fmt: x.number_format = sub_fmt
            W.row_dimensions[r].height = 18; W.row_dimensions[r + 1].height = 28; W.row_dimensions[r + 2].height = 18
        def section(r, text):
            W.merge_cells(start_row=r, start_column=2, end_row=r, end_column=13)
            x = W.cell(r, 2, text); x.font = F(bold=True, size=12, color=INK)
            for cc in range(2, 14): W.cell(r, cc).border = Border(bottom=Side(style='medium', color=CORAL))
            W.row_dimensions[r].height = 22

        section(4, 'O mês até agora')
        tile(5, 2, 'FATURAMENTO NO MÊS', f'={FAT}', BRL, f'="de "&TEXT({META},"R$ #,##0")&" de meta"', accent=True)
        tile(5, 4, '% DA META', f'=IFERROR({FAT}/{META},0)', PCT, f'="ritmo esperado: "&TEXT({pr("Dias úteis decorridos")}/{pr("Dias úteis no mês")},"0%")')
        tile(5, 6, 'FALTA PARA A META', f'=MAX(0,{META}-{FAT})', BRL, f'=TEXT({pr("Dias úteis restantes")},"0")&" dias úteis restantes"')
        tile(5, 8, 'PRECISA POR DIA', f'=IFERROR(MAX(0,{META}-{FAT})/{pr("Dias úteis restantes")},0)', BRL, f'="meta original: "&TEXT({META}/{pr("Dias úteis no mês")},"R$ #,##0")&"/dia"')
        tile(5, 10, 'PROJEÇÃO NO RITMO ATUAL', f'=IFERROR({FAT}/{pr("Dias úteis decorridos")}*{pr("Dias úteis no mês")},0)', BRL, f'=IF(IFERROR({FAT}/{pr("Dias úteis decorridos")}*{pr("Dias úteis no mês")},0)>={META},"✔ bate a meta",IF(IFERROR({FAT}/{pr("Dias úteis decorridos")}*{pr("Dias úteis no mês")},0)>={pr("Gatilho do bônus da equipe (R$)")},"libera bônus, não bate meta","✖ abaixo do gatilho de bônus"))')
        tile(5, 12, 'VS. MÊS ANTERIOR', f'=IFERROR({FAT}/{pr("Faturamento mês anterior até o mesmo dia (R$)")}-1,0)', '+0%;-0%;0%', f'="mesmo período: "&TEXT({pr("Faturamento mês anterior até o mesmo dia (R$)")},"R$ #,##0")')

        # dia de referência e próximo dia
        RU = pr('Último dia útil fechado')
        dref = f'SUMIFS({Bc("L")},{Bc("A")},{sg},{Bc("B")},{RU})'
        MD = f'{META}/{pr("Dias úteis no mês")}'
        alis_tm = f'IFERROR(SUMIFS({Bc("J")},{Bc("A")},{sg},{Bc("G")},"Alisamentos")/COUNTIFS({Bc("A")},{sg},{Bc("G")},"Alisamentos"),300)'
        atd = f'COUNTIFS({Ac("A")},{sg},{Ac("B")},{RU},{Ac("G")},"Finalizado")'
        if modo == 'tarde':
            # parcial do dia: o que entrou até agora e o que falta hoje
            need_hoje = f'IFERROR(MAX(0,{META}-({FAT}-{dref}))/({pr("Dias úteis restantes")}+1),0)'
            rest_ag = (f'(SUMIFS({Ac("H")},{Ac("A")},{sg},{Ac("B")},{RU})-SUMIFS({Ac("H")},{Ac("A")},{sg},{Ac("B")},{RU},{Ac("G")},"Finalizado")'
                       f'-SUMIFS({Ac("H")},{Ac("A")},{sg},{Ac("B")},{RU},{Ac("G")},"Cancelado")-SUMIFS({Ac("H")},{Ac("A")},{sg},{Ac("B")},{RU},{Ac("G")},"Cliente não compareceu"))')
            n_rest = (f'(COUNTIFS({Ac("A")},{sg},{Ac("B")},{RU})-COUNTIFS({Ac("A")},{sg},{Ac("B")},{RU},{Ac("G")},"Finalizado")'
                      f'-COUNTIFS({Ac("A")},{sg},{Ac("B")},{RU},{Ac("G")},"Cancelado")-COUNTIFS({Ac("A")},{sg},{Ac("B")},{RU},{Ac("G")},"Cliente não compareceu"))')
            section(9, f'Hoje até agora ({s["gerado"][11:16]}) → o que ainda falta hoje')
            tile(10, 2, f'FATURADO HOJE {refu:%d/%m}', f'={dref}', BRL, 'parcial, contas já fechadas', accent=True)
            tile(10, 4, 'META DE HOJE', f'={need_hoje}', BRL, 'para manter o mês na meta')
            tile(10, 6, 'FALTA HOJE', f'=MAX(0,{need_hoje}-{dref})', BRL, f'=IF({dref}>={need_hoje},"✔ meta de hoje batida","para fechar o dia na meta")', accent=True)
            tile(10, 8, 'ATENDIMENTOS HOJE', f'={atd}', INT, 'finalizados até agora')
            tile(10, 10, 'AINDA NA AGENDA HOJE', f'={rest_ag}', BRL, f'=TEXT({n_rest},"0")&" horários abertos"')
            tile(10, 12, 'FALTA DESCOBRIR HOJE', f'=MAX(0,{need_hoje}-{dref}-{rest_ag})', BRL, f'="≈ "&TEXT(ROUNDUP(MAX(0,{need_hoje}-{dref}-{rest_ag})/{alis_tm},0),"0")&" alisamentos de "&TEXT({alis_tm},"R$ #,##0")', accent=True)
        else:
            agd = f'SUMIFS(Agenda_{s["sig"]}!$G:$G,Agenda_{s["sig"]}!$A:$A,{PROX})'
            need = f'IFERROR(MAX(0,{META}-{FAT})/{pr("Dias úteis restantes")},0)'
            section(9, f'Fechamento de {refu:%d/%m}  →  plano para {prox:%d/%m}')
            tile(10, 2, f'FATURAMENTO {refu:%d/%m}', f'={dref}', BRL, f'="meta do dia: "&TEXT({MD},"R$ #,##0")')
            tile(10, 4, f'FALTOU EM {refu:%d/%m}', f'=MAX(0,{MD}-{dref})', BRL, f'=IF({dref}>={MD},"✔ meta do dia batida","diferença para a meta diária")')
            tile(10, 6, f'ATENDIMENTOS {refu:%d/%m}', f'={atd}', INT, f'="ticket médio: "&TEXT(IFERROR({dref}/COUNTIFS({Ac("A")},{sg},{Ac("B")},{RU},{Ac("I")},1),0),"R$ #,##0")')
            tile(10, 8, f'PRECISA VENDER {prox:%d/%m}', f'={need}', BRL, 'para fechar o mês na meta', accent=True)
            tile(10, 10, f'JÁ AGENDADO {prox:%d/%m}', f'={agd}', BRL, f'=TEXT(COUNTIFS(Agenda_{s["sig"]}!$A:$A,{PROX}),"0")&" horários na agenda"')
            tile(10, 12, f'FALTA DESCOBRIR {prox:%d/%m}', f'=MAX(0,{need}-{agd})', BRL, f'="≈ "&TEXT(ROUNDUP(MAX(0,{need}-{agd})/{alis_tm},0),"0")&" alisamentos de "&TEXT({alis_tm},"R$ #,##0")', accent=True)

        section(14, 'Clientes e equipe no mês')
        vis = f'COUNTIFS({Ac("A")},{sg},{Ac("I")},1)'
        tile(15, 2, 'CLIENTES NA LOJA', f'={vis}', INT, f'="meta "&{pr("Meta clientes na loja (recepção)")}&" · "&TEXT({vis}/{pr("Meta clientes na loja (recepção)")},"0%")')
        tile(15, 4, 'SAÍRAM COM RETORNO', f'=COUNTIFS({Ac("A")},{sg},{Ac("L")},1)', INT, f'="meta "&{pr("Meta clientes com retorno marcado")}&" · "&TEXT(COUNTIFS({Ac("A")},{sg},{Ac("L")},1)/{pr("Meta clientes com retorno marcado")},"0%")')
        tile(15, 6, 'VOLTARAM APÓS SUMIR', f'=COUNTIFS({Ac("A")},{sg},{Ac("K")},1)', INT, f'="base dos resgates · meta "&{pr("Meta resgates (closer da unidade)")}')
        tile(15, 8, 'NOVAS ASSINANTES', f'=COUNTIFS({Cc("A")},{sg},{Cc("B")},"Nova")', INT, f'=TEXT(COUNTIFS({Cc("A")},{sg},{Cc("B")},"Cancelada"),"0")&" canceladas · "&{pr("Assinaturas ativas hoje")}&" ativas"')
        tile(15, 10, 'CLIENTES NOVAS', f'={pr("Clientes novas cadastradas no mês")}', INT, f'=TEXT(COUNTIFS({Ac("A")},{sg},{Ac("J")},1),"0")&" já vieram e foram atendidas"')
        tile(15, 12, 'DESISTÊNCIAS', f'=COUNTIFS({Ac("A")},{sg},{Ac("G")},"Cancelado")+COUNTIFS({Ac("A")},{sg},{Ac("G")},"Cliente não compareceu")', INT,
             f'=TEXT(COUNTIFS(Resgate_{s["sig"]}!$F:$F,"Pendente – resgatar"),"0")&" pendentes de resgate"')

        # tabela por linha de receita (alimenta gráfico)
        section(19, 'Faturamento por linha de receita')
        hdr(W, 20, ['Linha', '', 'Realizado', 'Meta até hoje', 'Meta do mês', '% do mês'], 2)
        W.merge_cells('B20:C20')
        for i, (n, _) in enumerate(LINHAS):
            r = 21 + i
            W.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
            W.cell(r, 2, n)
            W.cell(r, 4, f'=SUMIFS({Bc("L")},{Bc("A")},{sg},{Bc("G")},B{r})')
            W.cell(r, 5, f'=F{r}*{pr("Dias úteis decorridos")}/{pr("Dias úteis no mês")}')
            W.cell(r, 6, f'=Premissas!$B${LR0+i}')
            W.cell(r, 7, f'=IFERROR(D{r}/F{r},0)')
            for cc, fm in zip(range(2, 8), [None, None, BRL, BRL, BRL, PCT]):
                x = W.cell(r, cc); x.font = F(size=10, color=INK); x.border = BOT
                if fm: x.number_format = fm
        rt = 21 + len(LINHAS)
        W.merge_cells(start_row=rt, start_column=2, end_row=rt, end_column=3)
        W.cell(rt, 2, 'Total')
        for cc, col in zip((4, 5, 6), 'DEF'): W.cell(rt, cc, f'=SUM({col}21:{col}{rt-1})')
        W.cell(rt, 7, f'=IFERROR(D{rt}/F{rt},0)')
        for cc, fm in zip(range(2, 8), [None, None, BRL, BRL, BRL, PCT]):
            x = W.cell(rt, cc); x.font = F(size=10, bold=True, color=INK); x.border = Border(top=Side(style='thin', color=INK))
            if fm: x.number_format = fm
        for rr in range(21, rt + 1): W.row_dimensions[rr].height = 30
        W.conditional_formatting.add(f'G21:G{rt-1}', DataBarRule(start_type='num', start_value=0, end_type='num', end_value=1, color=CORAL))
        # leitura rápida
        W.merge_cells(start_row=20, start_column=9, end_row=20, end_column=13)
        x = W.cell(20, 9, 'Leitura rápida'); x.font = F(bold=True, size=10, color=WHITE); x.fill = fill(INK); x.alignment = Alignment(indent=1, vertical='center')
        for k, txt in enumerate(s['insights']):
            W.merge_cells(start_row=21 + k, start_column=9, end_row=21 + k, end_column=13)
            y = W.cell(21 + k, 9, txt); y.font = F(size=9, color=INK); y.alignment = Alignment(wrap_text=True, vertical='top', indent=1)


        # gráficos
        section(rt + 2, 'Gráficos')
        g0 = rt + 3
        dcol = 3 + 5 * j
        cats = Reference(DI, min_col=2, min_row=D0, max_row=D1)
        bar = BarChart(); bar.type = 'col'; bar.title = 'Faturamento por dia útil × meta do dia'; bar.style = 2
        bar.add_data(Reference(DI, min_col=dcol, min_row=4, max_row=DREF), titles_from_data=True); bar.set_categories(cats)
        bar.series[0].graphicalProperties.solidFill = CORAL; bar.series[0].graphicalProperties.line.noFill = True
        bar.gapWidth = 40
        ln = LineChart(); ln.add_data(Reference(DI, min_col=dcol + 1, min_row=4, max_row=D1), titles_from_data=True)
        ln.series[0].graphicalProperties.line.solidFill = INK; ln.series[0].graphicalProperties.line.width = 19050
        ln.series[0].graphicalProperties.line.dashStyle = 'dash'; ln.series[0].smooth = False
        ln.y_axis.axId = 100; ln.y_axis.delete = True
        bar += ln
        bar.y_axis.numFmt = '"R$" #,##0'; bar.y_axis.majorGridlines.spPr = None
        bar.legend.position = 'b'; bar.height = 7.5; bar.width = 16
        bar.x_axis.delete = False; bar.y_axis.delete = False
        W.add_chart(bar, f'B{g0}')

        acc = LineChart(); acc.title = 'Acumulado do mês × meta acumulada'; acc.style = 2
        acc.add_data(Reference(DI, min_col=dcol + 2, min_row=4, max_row=DREF), titles_from_data=True)
        acc.add_data(Reference(DI, min_col=dcol + 3, min_row=4, max_row=D1), titles_from_data=True)
        acc.set_categories(cats)
        acc.series[0].graphicalProperties.line.solidFill = CORAL; acc.series[0].graphicalProperties.line.width = 28575
        acc.series[1].graphicalProperties.line.solidFill = INK; acc.series[1].graphicalProperties.line.dashStyle = 'dash'; acc.series[1].graphicalProperties.line.width = 19050
        for se in acc.series: se.smooth = False; se.marker.symbol = 'none'
        acc.y_axis.numFmt = '"R$" #,##0'; acc.legend.position = 'b'; acc.height = 7.5; acc.width = 16
        acc.x_axis.delete = False; acc.y_axis.delete = False
        W.add_chart(acc, f'H{g0}')

        lr = BarChart(); lr.type = 'bar'; lr.title = 'Linhas de receita: realizado × meta até hoje'; lr.style = 2
        lr.add_data(Reference(W, min_col=4, min_row=20, max_row=rt - 1), titles_from_data=True)
        lr.add_data(Reference(W, min_col=5, min_row=20, max_row=rt - 1), titles_from_data=True)
        lr.set_categories(Reference(W, min_col=2, min_row=21, max_row=rt - 1))
        lr.series[0].graphicalProperties.solidFill = CORAL; lr.series[1].graphicalProperties.solidFill = 'C9C6D0'
        for se in lr.series: se.graphicalProperties.line.noFill = True
        lr.x_axis.scaling.orientation = 'maxMin'; lr.y_axis.numFmt = '"R$" #,##0'; lr.gapWidth = 60; lr.overlap = -10
        lr.legend.position = 'b'; lr.height = 7.5; lr.width = 16; lr.x_axis.delete = False; lr.y_axis.delete = False
        W.add_chart(lr, f'B{g0+16}')

        # ---- Equipe
        E = wb.create_sheet(f'Equipe {s["sig"]}', j * 3 + 1)
        title(E, f'Equipe · {s["loja"]}', f'Desempenho de cada colaboradora no mês, até {ref:%d/%m}. Uma linha por pessoa.', 14)
        widths(E, [22, 12, 12, 13, 10, 12, 13, 12, 11, 27, 10, 11, 10, 12])
        eh = ['Colaboradora', 'Função', 'Atendimentos', 'Faturamento serviços', 'Produtos (un.)', 'Produtos (R$)', 'Total faturado',
              'Assinaturas vendidas', 'Contas ≥ R$500', 'Meta principal', 'Meta', 'Realizado', '% da meta', 'Precisa por dia']
        hdr(E, 4, eh)
        cfg = s['cfg']
        ordem = [v for v in cfg.values() if v[3] is not None] + [v for v in cfg.values() if v[3] is None]
        nomes_cfg = [v[0] for v in ordem]
        NAO_PESSOA = {'Sem profissional', 'Clube (cobrança automática)', 'Site / online', 'Não informado'}
        extras = [p for p in s['pessoas'] if p not in nomes_cfg and p not in NAO_PESSOA]
        lista = [(v[0], v[1], v[2], v[3], v[4]) for v in ordem] + [(p, '–', None, None, None) for p in extras]
        ativos = {i[7] for i in s['itens']} | {a[5] for a in s['agenda']} | {c[7] for c in s['clube']}
        lista = [x for x in lista if x[3] is not None or x[0] in ativos]
        for i, (nome, func, mlab, mval, mtype) in enumerate(lista):
            r = 5 + i; q = f'$A{r}'
            real = {'visitas': f'=COUNTIFS({Ac("A")},{sg},{Ac("I")},1)', 'resgates': f'=COUNTIFS({Ac("A")},{sg},{Ac("K")},1)',
                    'atend': f'=C{r}', 'fat': f'=D{r}'}.get(mtype)
            vals = [nome, func,
                    f'=COUNTIFS({Ac("A")},{sg},{Ac("F")},{q},{Ac("G")},"Finalizado")',
                    f'=SUMIFS({Bc("L")},{Bc("A")},{sg},{Bc("H")},{q},{Bc("E")},"Serviço")',
                    f'=SUMIFS({Bc("I")},{Bc("A")},{sg},{Bc("H")},{q},{Bc("E")},"Produto")',
                    f'=SUMIFS({Bc("L")},{Bc("A")},{sg},{Bc("H")},{q},{Bc("E")},"Produto")',
                    f'=D{r}+F{r}',
                    f'=COUNTIFS({Cc("A")},{sg},{Cc("H")},{q},{Cc("B")},"Nova")',
                    f'=SUMIFS({Bc("N")},{Bc("A")},{sg},{Bc("H")},{q})',
                    mlab or '–', mval, real,
                    f'=IFERROR(L{r}/K{r},"")' if real else None,
                    f'=IFERROR(MAX(0,K{r}-L{r})/{pr("Dias úteis restantes")},"")' if real else None]
            money = mtype == 'fat'
            fmts = {3: INT, 4: BRL, 5: INT, 6: BRL, 7: BRL, 8: INT, 9: INT, 11: BRL if money else INT, 12: BRL if money else INT,
                    13: PCT, 14: BRL if money else '0.0'}
            for c, v in enumerate(vals, 1):
                x = E.cell(r, c, v); x.font = F(size=10, color=INK, bold=(c in (1, 7))); x.border = BOT
                x.number_format = fmts.get(c, 'General')
                if c >= 3 and c != 10: x.alignment = Alignment(horizontal='right')
                if i % 2: x.fill = fill(BG)
            E.row_dimensions[r].height = 20
        nE = 4 + len(lista)
        E.conditional_formatting.add(f'M5:M{nE}', DataBarRule(start_type='num', start_value=0, end_type='num', end_value=1, color=CORAL))
        rn = nE + 2
        notes = ['Atendimentos = horários finalizados na agenda da Trinks. Faturamento já desconta benefícios usados por assinantes do clube.',
                 'Contas ≥ R$500 = contas fechadas em R$ 500 ou mais em que a pessoa fez algum serviço (base do PIX de R$ 25 das especialistas).',
                 'A Trinks não registra quem resgatou cada cliente: o realizado da closer é o número da loja de clientes que voltaram após 90+ dias ou após faltar.',
                 'Recepção: o realizado é o total de clientes atendidas na loja. Na Trinks: Daiane = Encantadora Aux. 1, Julia = Aux. 2, Rosana = Aux. 3, Laiane = Manicure.']
        for k, t in enumerate(notes):
            E.cell(rn + k, 1, t).font = F(size=9, color=MUTED, italic=True)
        def lbl(ch, fmt=None):
            ch.dataLabels = DataLabelList(); ch.dataLabels.showVal = True
            ch.dataLabels.showSerName = False; ch.dataLabels.showCatName = False; ch.dataLabels.showLegendKey = False
            if fmt: ch.dataLabels.numFmt = fmt
        ch = BarChart(); ch.type = 'bar'; ch.title = 'Total faturado por colaboradora'; ch.style = 2
        ch.add_data(Reference(E, min_col=7, min_row=5, max_row=nE), titles_from_data=False)
        ch.set_categories(Reference(E, min_col=1, min_row=5, max_row=nE))
        ch.series[0].graphicalProperties.solidFill = CORAL; ch.series[0].graphicalProperties.line.noFill = True
        ch.legend = None; ch.x_axis.scaling.orientation = 'maxMin'; ch.y_axis.numFmt = '"R$" #,##0'; ch.gapWidth = 50
        ch.height = 8; ch.width = 15; ch.x_axis.delete = False; ch.y_axis.delete = False; lbl(ch, '"R$" #,##0')
        E.add_chart(ch, f'A{rn + len(notes) + 1}')
        ch2 = BarChart(); ch2.type = 'bar'; ch2.title = 'Atendimentos finalizados por colaboradora'; ch2.style = 2
        ch2.add_data(Reference(E, min_col=3, min_row=5, max_row=nE), titles_from_data=False)
        ch2.set_categories(Reference(E, min_col=1, min_row=5, max_row=nE))
        ch2.series[0].graphicalProperties.solidFill = INK2; ch2.series[0].graphicalProperties.line.noFill = True
        ch2.legend = None; ch2.x_axis.scaling.orientation = 'maxMin'; ch2.gapWidth = 50
        ch2.height = 8; ch2.width = 15; ch2.x_axis.delete = False; ch2.y_axis.delete = False; lbl(ch2)
        E.add_chart(ch2, f'H{rn + len(notes) + 1}')
        E.freeze_panes = 'B5'

        # ---- Agenda do próximo dia e futuros
        A = wb.create_sheet(f'Agenda_{s["sig"]}', j * 3 + 2)
        title(A, f'Agenda · {s["loja"]}', f'Horários marcados a partir de {prox:%d/%m} (sem cancelados). Os de {prox:%d/%m} entram no cálculo do Painel.', 8)
        nfa = table(A, 4, ['Data', 'Hora', 'Cliente', 'Serviço', 'Profissional', 'Status', 'Valor'], [f[1:] for f in s['futuro']], ['dd/mm/yyyy', None, None, None, None, None, BRL])
        for r in range(5, nfa + 1):
            if A.cell(r, 1).value == prox:
                for c in range(1, 8): A.cell(r, c).fill = fill(CORAL_L)
        widths(A, [12, 7, 32, 26, 24, 14, 11])
        if not s['futuro']: A.cell(5, 1, 'Nenhum horário marcado.').font = F(size=10, color=MUTED)

        # ---- Resgate (lista de ação para closers)
        R = wb.create_sheet(f'Resgate_{s["sig"]}')
        title(R, f'Lista de resgate · {s["loja"]}', 'Para a closer da unidade: quem marcou e não veio no mês, e quem está há 60+ dias sem vir e sem horário marcado.', 6)
        widths(R, [12, 32, 24, 24, 24, 24])
        R.cell(4, 1, 'Marcaram e desistiram no mês').font = F(bold=True, size=11, color=INK)
        n1 = table(R, 5, ['Data', 'Cliente', 'Serviço', 'Profissional', 'Status', 'Situação'], s['desist'], ['dd/mm/yyyy', None, None, None, None, None])
        for r in range(6, n1 + 1):
            if R.cell(r, 6).value == 'Pendente – resgatar': R.cell(r, 6).font = F(size=10, bold=True, color=RED)
            elif str(R.cell(r, 6).value).startswith('Voltou'): R.cell(r, 6).font = F(size=10, color=GREEN)
        r2 = n1 + 3
        R.cell(r2, 1, f'Há 60+ dias sem vir e sem horário marcado ({len(s["sumidas"])} clientes, das mais recentes para as mais antigas)').font = F(bold=True, size=11, color=INK)
        table(R, r2 + 1, ['Última visita', 'Cliente', 'Dias sem vir', 'Último serviço', 'Profissional', 'Visitas no histórico'],
              [[x[1], x[0], x[2], x[3], x[4], x[5]] for x in s['sumidas']], ['dd/mm/yyyy', None, INT, None, None, INT])
        R.cell(r2 + 2 + len(s['sumidas']), 1, 'Histórico considerado: agenda da Trinks desde o início do 4º mês anterior.').font = F(size=9, color=MUTED, italic=True)

        # ---- Semanas (visão semanal para a reunião de segunda)
        S = wb.create_sheet(f'Semanas {s["sig"]}')
        title(S, f'Semana a semana · {s["loja"]}', 'Semanas de segunda a domingo dentro do mês. Meta da semana = meta do dia × dias úteis da semana.', 12)
        widths(S, [22, 12, 12, 12, 12, 10, 12, 12, 12, 12, 12, 12])
        semanas = []; d = dias_mes[0]
        while d.month == ref.month:
            fim = d + dt.timedelta(days=6 - d.weekday())
            if fim.month != ref.month: fim = dias_mes[-1]
            semanas.append((d, fim)); d = fim + dt.timedelta(days=1)
        hdr(S, 4, ['Semana', 'Início', 'Fim', 'Dias úteis', 'Meta da semana', 'Realizado', '% da meta', 'Atendimentos', 'Clientes na loja', 'Voltaram após sumir', 'Novas assinantes', 'Situação'])
        for i, (a, b) in enumerate(semanas):
            r = 5 + i
            rng = lambda col, base: f'{base(col)},">="&$B{r},{base(col)},"<="&$C{r}'
            vals = [f'Semana {i+1}', a, b,
                    f'=COUNTIFS(Diário!$A${D0}:$A${D1},">="&B{r},Diário!$A${D0}:$A${D1},"<="&C{r})',
                    f'=D{r}*{MD}',
                    f'=IF(B{r}>{REF},"",SUMIFS({Bc("L")},{Bc("A")},{sg},{rng("B", Bc)}))',
                    f'=IF(F{r}="","",IFERROR(F{r}/E{r},0))',
                    f'=IF(B{r}>{REF},"",COUNTIFS({Ac("A")},{sg},{Ac("G")},"Finalizado",{rng("B", Ac)}))',
                    f'=IF(B{r}>{REF},"",COUNTIFS({Ac("A")},{sg},{Ac("I")},1,{rng("B", Ac)}))',
                    f'=IF(B{r}>{REF},"",COUNTIFS({Ac("A")},{sg},{Ac("K")},1,{rng("B", Ac)}))',
                    f'=IF(B{r}>{REF},"",COUNTIFS({Cc("A")},{sg},{Cc("B")},"Nova",{rng("C", Cc)}))',
                    f'=IF(B{r}>{REF},"futura",IF(C{r}>{REF},"em andamento",IF(F{r}>=E{r},"✔ bateu","✖ abaixo")))']
            for c, v in enumerate(vals, 1):
                x = S.cell(r, c, v); x.font = F(size=10, color=INK, bold=(c == 1)); x.border = BOT
                x.number_format = {2: 'dd/mm', 3: 'dd/mm', 4: INT, 5: BRL, 6: BRL, 7: PCT, 8: INT, 9: INT, 10: INT, 11: INT}.get(c, 'General')
                if a <= refu <= b: x.fill = fill(CORAL_L)
        nS = 4 + len(semanas)
        S.conditional_formatting.add(f'G5:G{nS}', DataBarRule(start_type='num', start_value=0, end_type='num', end_value=1, color=CORAL))
        wc = BarChart(); wc.type = 'col'; wc.title = 'Realizado × meta por semana'; wc.style = 2
        wc.add_data(Reference(S, min_col=6, min_row=4, max_row=nS), titles_from_data=True)
        wc.add_data(Reference(S, min_col=5, min_row=4, max_row=nS), titles_from_data=True)
        wc.set_categories(Reference(S, min_col=1, min_row=5, max_row=nS))
        wc.series[0].graphicalProperties.solidFill = CORAL; wc.series[1].graphicalProperties.solidFill = 'C9C6D0'
        for se in wc.series: se.graphicalProperties.line.noFill = True
        wc.y_axis.numFmt = '"R$" #,##0'; wc.legend.position = 'b'; wc.height = 7.5; wc.width = 16; wc.gapWidth = 60
        wc.x_axis.delete = False; wc.y_axis.delete = False
        S.add_chart(wc, f'A{nS + 2}')
        # colaboradoras por semana
        r0 = nS + 19
        S.cell(r0, 1, 'Colaboradoras por semana').font = F(bold=True, size=11, color=INK)
        heads = ['Colaboradora']
        for i in range(len(semanas)): heads += [f'S{i+1} atend.', f'S{i+1} R$']
        hdr(S, r0 + 1, heads)
        pessoas_sem = [x[0] for x in lista]
        for k, nome in enumerate(pessoas_sem):
            r = r0 + 2 + k
            S.cell(r, 1, nome).font = F(size=10, bold=True, color=INK)
            for i in range(len(semanas)):
                ca, cf = 2 + 2 * i, 3 + 2 * i
                B_, C_ = f'$B${5+i}', f'$C${5+i}'
                S.cell(r, ca, f'=IF({B_}>{REF},"",COUNTIFS({Ac("A")},{sg},{Ac("F")},$A{r},{Ac("G")},"Finalizado",{Ac("B")},">="&{B_},{Ac("B")},"<="&{C_}))').number_format = INT
                S.cell(r, cf, f'=IF({B_}>{REF},"",SUMIFS({Bc("L")},{Bc("A")},{sg},{Bc("H")},$A{r},{Bc("B")},">="&{B_},{Bc("B")},"<="&{C_}))').number_format = BRL
            for c in range(1, len(heads) + 1):
                S.cell(r, c).border = BOT
                if k % 2: S.cell(r, c).fill = fill(BG)

    # ---- Caixa (todas as lojas)
    K = wb.create_sheet('Caixa')
    formas = sorted({p[4] for p in pags})
    title(K, 'Caixa por forma de pagamento', f'Fechamento de {ref:%d/%m} no topo; abaixo, o mês dia a dia. A Trinks não informa quem fechou o caixa.', 3 + len(formas))
    widths(K, [8, 10] + [14] * len(formas) + [14])
    r = 4
    for s in stores:
        sg = f'"{s["sig"]}"'
        K.cell(r, 1, f'{s["loja"]} · fechamento de {ref:%d/%m}').font = F(bold=True, size=11, color=INK)
        hdr(K, r + 1, ['Loja', 'Dia'] + formas + ['Total'])
        rows_d = [ref] + [d for d in dias_mes if d <= ref]
        for k, d in enumerate(rows_d):
            rr = r + 2 + k + (1 if k > 0 else 0)
            if k == 1:
                K.cell(rr - 1, 1, 'Mês dia a dia').font = F(bold=True, size=10, color=INK2)
            K.cell(rr, 1, s['sig']); K.cell(rr, 2, d).number_format = 'dd/mm'
            for fi, fn in enumerate(formas):
                K.cell(rr, 3 + fi, f'=SUMIFS({Pc("F")},{Pc("A")},$A{rr},{Pc("B")},$B{rr},{Pc("E")},"{fn}")').number_format = BRL2
            K.cell(rr, 3 + len(formas), f'=SUM(C{rr}:{L(2+len(formas))}{rr})').number_format = BRL2
            for c in range(1, 4 + len(formas)):
                x = K.cell(rr, c); x.border = BOT; x.font = F(size=10, color=INK, bold=(k == 0))
                if k == 0: x.fill = fill(CORAL_L)
        last = r + 2 + len(rows_d)
        K.cell(last + 1, 2, 'Total mês').font = F(bold=True, size=10)
        for fi in range(len(formas) + 1):
            col = L(3 + fi)
            c = K.cell(last + 1, 3 + fi, f'=SUM({col}{r+4}:{col}{last})'); c.number_format = BRL2; c.font = F(bold=True, size=10)
            c.border = Border(top=Side(style='thin', color=INK))
        r = last + 4
    K.freeze_panes = 'C4'

    # ordem final das abas
    order = []
    for s in stores:
        o = [f'Painel {s["sig"]}', f'Equipe {s["sig"]}', f'Semanas {s["sig"]}', f'Resgate_{s["sig"]}', f'Agenda_{s["sig"]}']
        if semanal: o.insert(1, o.pop(2))      # segunda-feira: semana logo após o painel
        order += o
    order += ['Caixa', 'Diário', 'Premissas', 'Base_Receita', 'Base_Pagamentos', 'Base_Agenda', 'Base_Clube']
    wb._sheets = [wb[n] for n in order]
    for ws in wb.worksheets:
        ws.sheet_properties.tabColor = CORAL if ws.title.startswith('Painel') else (INK if not ws.title.startswith('Base') else 'C9C6D0')
        if not ws.title.startswith('Base'):
            ws.page_setup.orientation = 'landscape'; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
            ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.active = 0
    wb.save(out)

def insights(s):
    """Frases curtas calculadas a partir dos dados (aparecem no painel)."""
    ref = s['ref']; out = []
    R = lambda v: 'R$ ' + f'{v:,.0f}'.replace(',', '.')
    it = s['itens']
    fat = sum(i[11] for i in it)
    dias = [d for d in workdays(ref.year, ref.month) if is_util(d)]
    dec = sum(1 for d in dias if d <= ref); rest = len(dias) - dec
    meta = META_MES.get(f'{ref:%Y-%m}', 50000)
    if dec: out.append(f'Média de {R(fat/dec)} por dia útil. Para bater a meta, os próximos {rest} dias úteis precisam de {R(max(0, meta-fat)/max(rest,1))} cada.')
    al = [i for i in it if i[6] == 'Alisamentos']
    if al:
        tm = sum(i[9] for i in al) / len(al)
        out.append(f'{len(al)} alisamentos no mês, com ticket médio de {R(tm)} (o plano considera R$ 400).')
    club = sum(i[11] for i in it if i[6] == 'Assinatura Encanto')
    out.append(f'Clube: {s["ativas"]} assinaturas ativas, cerca de {R(s["ativas"]*200)}/mês, contra R$ 8.000 previstos no plano.')
    pend = sum(1 for d in s['desist'] if d[5].startswith('Pendente'))
    if pend: out.append(f'{pend} clientes que desmarcaram ou faltaram ainda não voltaram: estão na aba Resgate.')
    return out[:4]

if __name__ == '__main__':
    out = sys.argv[1]
    stores = []
    for p in sys.argv[2:]:
        s = prepare(load(p)); s['insights'] = insights(s); stores.append(s)
    build(out, stores)
    print('ok', out)

