"""
Диагностический скрипт — запусти и скинь вывод.
Показывает точные цифры по подсетям и хостам.
"""
import ipaddress
import collections
import pandas as pd

SUBNET_FILE = "subnet.xlsx"
HOST_FILE   = "host.xlsx"
SUBNET_COL  = "subnet"
HOST_COL    = "host"


def load_column(filepath, column):
    df = pd.read_excel(filepath, dtype=str)
    df.columns = df.columns.str.strip()
    col_map = {c.lower(): c for c in df.columns}
    matched = col_map.get(column.lower())
    if not matched:
        raise ValueError("Колонка '%s' не найдена. Доступные: %s" % (column, list(df.columns)))
    return df[matched].dropna().str.strip().tolist()


def normalise(s):
    try:
        return str(ipaddress.ip_network(s, strict=False))
    except ValueError:
        return None


# ── Подсети ───────────────────────────────────────────────────────────────────
raw_subnets = load_column(SUBNET_FILE, SUBNET_COL)
print("=== ПОДСЕТИ ===")
print("Строк в subnet.xlsx (raw):          %d" % len(raw_subnets))

norm_subnets = []
bad_subnets  = []
for s in raw_subnets:
    n = normalise(s)
    if n:
        norm_subnets.append(n)
    else:
        bad_subnets.append(s)

print("Успешно нормализовано:              %d" % len(norm_subnets))
print("Не удалось разобрать (bad):         %d" % len(bad_subnets))
if bad_subnets:
    print("  Примеры bad: %s" % bad_subnets[:5])

deduped = list(dict.fromkeys(norm_subnets))   # дедупликация с сохранением порядка
print("После дедупликации (уникальных):   %d" % len(deduped))
dups = len(norm_subnets) - len(deduped)
print("Дублей удалено:                    %d" % dups)

# ── Хосты ─────────────────────────────────────────────────────────────────────
raw_hosts = load_column(HOST_FILE, HOST_COL)
print("\n=== ХОСТЫ ===")
print("Строк в host.xlsx (raw):            %d" % len(raw_hosts))

networks = []
for s in deduped:
    try:
        networks.append(ipaddress.ip_network(s, strict=False))
    except ValueError:
        pass

matched_subnets = set()
unmatched_hosts = []
for host in raw_hosts:
    try:
        ip = ipaddress.ip_address(host)
    except ValueError:
        unmatched_hosts.append(host)
        continue
    hits = [net for net in networks if ip in net]
    if hits:
        best = str(max(hits, key=lambda n: n.prefixlen))
        matched_subnets.add(best)
    else:
        unmatched_hosts.append(host)

print("Хостов с найденной подсетью:        %d" % (len(raw_hosts) - len(unmatched_hosts)))
print("Хостов без подсети (уникальные):    %d" % len(unmatched_hosts))

# ── Подсети без хостов ────────────────────────────────────────────────────────
unique_subnets = [s for s in deduped if s not in matched_subnets]
print("\n=== ИТОГ ===")
print("Всего уникальных подсетей:          %d" % len(deduped))
print("Подсетей с хостами (matched):       %d" % len(matched_subnets))
print("Подсетей БЕЗ хостов (уникальные):  %d" % len(unique_subnets))

if unique_subnets:
    print("\nПервые 10 уникальных подсетей:")
    for s in unique_subnets[:10]:
        print("  %s" % s)
