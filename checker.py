import pandas as pd
import requests
from concurrent.futures import ThreadPoolExecutor
import argparse

def check_url(url, timeout):
    if not url.startswith(('http://', 'https://')):
        url = 'https://' + url
    
    try:
        response = requests.get(url, timeout=timeout, allow_redirects=True)
        status = response.status_code
        final_url = response.url
        redirect_count = len(response.history)
        
        title = ''
        if '<title>' in response.text:
            start = response.text.find('<title>') + 7
            end = response.text.find('</title>')
            title = response.text[start:end][:80]
        
        return [url, status, title, final_url, redirect_count, '']
        
    except requests.exceptions.Timeout:
        return [url, 'ошибка', '', '', 0, 'таймаут']
    except Exception as e:
        return [url, 'ошибка', '', '', 0, str(e)[:40]]

parser = argparse.ArgumentParser(description='Массовая проверка сайтов')
parser.add_argument('input', nargs='?', default='links.xlsx', help='Входной Excel файл')
parser.add_argument('output', nargs='?', default='result', help='Имя выходного файла')
parser.add_argument('--timeout', type=int, default=10, help='Таймаут в секундах')
parser.add_argument('--workers', type=int, default=10, help='Количество потоков')
args = parser.parse_args()

print(f"Читаю файл: {args.input}")
print(f"Таймаут: {args.timeout} сек, Потоков: {args.workers}")

df = pd.read_excel(args.input)
print(f"Найдено ссылок: {len(df)}")

with ThreadPoolExecutor(max_workers=args.workers) as executor:
    results = list(executor.map(lambda url: check_url(url, args.timeout), df['url']))

result_df = pd.DataFrame(results, columns=['исходный_url', 'статус', 'тайтл', 'финальный_url', 'число_редиректов', 'ошибка'])

result_df.to_excel(f'{args.output}.xlsx', index=False)
result_df.to_csv(f'{args.output}.csv', index=False, encoding='utf-8-sig')

working = len(result_df[result_df['статус'] == 200])
broken = len(result_df[result_df['статус'] == 404])
errors = len(result_df[result_df['статус'] == 'ошибка'])
redirects = len(result_df[result_df['число_редиректов'] > 0])

print("\nСтатистика:")
print(f"Работают: {working}")
print(f"Битые (404): {broken}")
print(f"Ошибки: {errors}")
print(f"С редиректами: {redirects}")
print(f"Результат: {args.output}.xlsx")

broken_df = result_df[result_df['статус'].isin([404, 'ошибка'])]
if len(broken_df) > 0:
    broken_df.to_excel(f'{args.output}_broken.xlsx', index=False)
    print(f"Битые ссылки сохранены в {args.output}_broken.xlsx")