#!/usr/bin/env python3
"""Onde comprar minério sem concorrente local — cobertura por ALCANCE da ordem de compra.

Uso:
    python3 sistemas_sem_comprador.py --regiao 10000002 --hub 30000142 [--belts]

Regiões: The Forge 10000002 (Jita) · Domain 10000043 (Amarr) · Sinq Laison 10000032
(Dodixie) · Metropolis 10000042 (Hek) · Heimatar 10000030 (Rens).

Por que a pergunta "tem buy order no sistema?" dá resposta errada: ordem de compra tem
campo `range` (station/solarsystem/N saltos/region). Ordem com range de região deixa
qualquer um vender de qualquer estação da região — o sistema parece "sem comprador"
olhando a estação dele, e não está. A métrica que vale é o preço MÁXIMO ALCANÇÁVEL
sem sair da estação, comparado ao preço executável no hub.

SDE leve vem do dump da Fuzzwork (~/Documentos/Projetos/EveTrade/eve_data/ para
invTypes.csv; os map* ficam em cache local). mapDenormalize.csv (83 MB) só é baixado
com --belts e serve para contar cinturões por sistema.
"""
import argparse
import collections
import csv
import json
import os
import statistics
import urllib.request

CACHE = os.path.expanduser('~/Documentos/Projetos/EveTrade/eve_data/sde')
DUMP = 'https://www.fuzzwork.co.uk/dump/latest/csv/'
TIPOS = {  # nome -> (cru, comprimido, m3 por unidade crua)
    'Veldspar': (1230, 62516, 0.1), 'Scordite': (1228, 62520, 0.15),
    'Pyroxeres': (1224, 62524, 0.3), 'Plagioclase': (18, 62528, 0.35),
    'Omber': (178, 62532, 0.6), 'Kernite': (20, 62536, 1.2),
    'Jaspet': (1226, 62540, 2.0), 'Hemorphite': (1231, 62544, 3.0),
    'Hedbergite': (21, 62548, 3.0), 'Gneiss': (1229, 62552, 5.0),
    'Dark Ochre': (1232, 62556, 8.0), 'Crokite': (1227, 62560, 16.0),
    'Arkonor': (22, 62568, 16.0), 'Bistot': (1223, 62564, 16.0),
    'Spodumain': (19, 62572, 16.0), 'Mercoxit': (1132, 62586, 40.0),
}


def baixar(nome, forcar=False):
    os.makedirs(CACHE, exist_ok=True)
    p = os.path.join(CACHE, nome + '.csv')
    if not os.path.exists(p) or forcar:
        urllib.request.urlretrieve(DUMP + nome + '.csv', p)
    return p


def ordens_regiao(regiao, tipo):
    out = []
    for p in range(1, 12):
        u = (f'https://esi.evetech.net/latest/markets/{regiao}/orders/?datasource=tranquility'
             f'&order_type=buy&type_id={tipo}&page={p}')
        try:
            d = json.load(urllib.request.urlopen(u, timeout=60))
        except Exception:
            break
        out += d
        if len(d) < 1000:
            break
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--regiao', type=int, default=10000002)
    ap.add_argument('--hub', type=int, default=30000142)
    ap.add_argument('--belts', action='store_true')
    ap.add_argument('--min-volume', type=int, default=1000, help='descarta ordem-isca')
    a = ap.parse_args()

    sysm = {}
    for r in csv.DictReader(open(baixar('mapSolarSystems'), encoding='utf-8-sig')):
        sysm[int(r['solarSystemID'])] = {'nome': r['solarSystemName'],
                                         'sec': float(r['security']),
                                         'reg': int(r['regionID'])}
    adj = collections.defaultdict(set)
    for r in csv.DictReader(open(baixar('mapSolarSystemJumps'), encoding='utf-8-sig')):
        x, y = int(r['fromSolarSystemID']), int(r['toSolarSystemID'])
        adj[x].add(y)
        adj[y].add(x)
    sta = {int(r['stationID']): int(r['solarSystemID'])
           for r in csv.DictReader(open(baixar('staStations'), encoding='utf-8-sig'))}

    def bfs(o, lim=None):
        d = {o: 0}
        q = collections.deque([o])
        while q:
            u = q.popleft()
            if lim is not None and d[u] >= lim:
                continue
            for v in adj[u]:
                if v not in d:
                    d[v] = d[u] + 1
                    q.append(v)
        return d
    dh = bfs(a.hub)

    belts = collections.Counter()
    if a.belts:
        for r in csv.DictReader(open(baixar('mapDenormalize'), encoding='utf-8-sig')):
            if r['groupID'] == '9':
                try:
                    belts[int(r['solarSystemID'])] += 1
                except ValueError:
                    pass

    cob = collections.defaultdict(dict)
    for nome, (cru, comp, v) in TIPOS.items():
        for tid, vol_un in ((cru, v), (comp, v)):
            for o in ordens_regiao(a.regiao, tid):
                if o['volume_remain'] < a.min_volume:
                    continue
                s = sta.get(o['location_id'])
                if not s or sysm.get(s, {}).get('reg') != a.regiao:
                    continue
                m3 = o['price'] / vol_un
                rng = o['range']
                if rng == 'region':
                    alvo = [k for k in sysm if sysm[k]['reg'] == a.regiao]
                elif rng in ('station', 'solarsystem'):
                    alvo = [s]
                else:
                    alvo = bfs(s, int(rng))
                for sid in alvo:
                    if sysm[sid]['reg'] != a.regiao:
                        continue
                    if m3 > cob[sid].get(nome, (0, 0))[0]:
                        cob[sid][nome] = (m3, o['volume_remain'])

    hs = [k for k, v in sysm.items() if v['reg'] == a.regiao and v['sec'] >= 0.45]
    if a.belts:
        hs = [k for k in hs if belts.get(k, 0) >= 3]
    print(f'{len(hs)} sistemas highsec na região {a.regiao}')
    print(f"{'minério':<12}{'remota med':>11}{'sist. sem oferta real':>23}")
    for nome in TIPOS:
        vals = [cob[str(k)].get(nome, (0, 0))[0] for k in hs]
        med = statistics.median(vals) if vals else 0
        vazios = sum(1 for x in vals if x < 30)
        print(f'{nome:<12}{med:>11.1f}{vazios:>16}/{len(hs):<6}')

    print('\nSistemas com cinturão, oferta remota fraca e perto do hub:')
    cand = [(k, max(cob[str(k)].get(n, (0, 0))[0] for n in ('Veldspar', 'Scordite')))
            for k in hs]
    cand = [(k, p) for k, p in cand if p < 70 and k in dh]
    for k, p in sorted(cand, key=lambda x: (dh[x[0]], -belts.get(x[0], 0)))[:20]:
        print(f"  {sysm[k]['nome']:<18}{dh[k]:>2} saltos  {belts.get(k, 0):>3} cinturões  "
              f"oferta {p:>5.0f} ISK/m3")


if __name__ == '__main__':
    main()
