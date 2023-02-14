import json
from datetime import datetime

import requests
import xlsxwriter

cookies = {
    'yuidss': '2928503961649669498',
    'yandexuid': '2928503961649669498',
    'font_loaded': 'YSv1',
    'my': 'YwA=',
    'gdpr': '0',
    'skid': '7515762271662945052',
    '_ym_d': '1663497667',
    'ymex': '1965385734.yrts.1650025734#1982640693.yrtsi.1667280693',
    '_ym_uid': '16700433131044207680',
    'i': 'c6SRcyO55mhKKG63GE+jdup+alyhqo9fAOZ2xkWTNNUtTrHubzpRlGh53kmTwjneWcBLLhEvSJr7enmJtgm+gqkbUrw=',
    'ys': 'vbch.2-35-0',
    'yashr': '6774375601676365280',
    'PHPSESSID': 'b958908d589c4751a619b529e1c1e461',
    'SL_G_WPT_TO': 'ru',
    'SL_GWPT_Show_Hide_tmp': '1',
    'SL_wptGlobTipTmp': '1',
    'is_gdpr': '0',
    'is_gdpr_b': 'CKWxOxDNpgE=',
    '_yasc': 'IB2IJ3hkPKKs2lEqh2BZ4Rsk6eRqHOxSMhjHpjzKjXmp1e3xx4LRgi9QccB0ipw=',
    'eda_web': '%7B%22app%22%3A%7B%22lat%22%3A55.755245%2C%22lon%22%3A37.617779%2C%22deliveryTime%22%3Anull%2C%22xDeviceId%22%3A%22l6lo13q2-melqfb568r-57k4dzq0n59-l9q1wy8nfze%22%2C%22appBannerShown%22%3Afalse%2C%22isAdult%22%3Anull%2C%22yandexPlusCashbackOptInChecked%22%3Afalse%2C%22testRunId%22%3Anull%2C%22initialPromocode%22%3Anull%2C%22themeVariantKey%22%3A%22light%22%2C%22lang%22%3A%22ru%22%2C%22translateMenu%22%3Afalse%7D%7D',
}
headers = {
    'authority': 'eda.yandex.ru',
    'accept': 'application/json, text/plain, */*',
    'accept-language': 'ru',
    'content-type': 'application/json;charset=UTF-8',
    # 'cookie': 'yuidss=2928503961649669498; yandexuid=2928503961649669498; font_loaded=YSv1; my=YwA=; gdpr=0; skid=7515762271662945052; _ym_d=1663497667; ymex=1965385734.yrts.1650025734#1982640693.yrtsi.1667280693; _ym_uid=16700433131044207680; i=c6SRcyO55mhKKG63GE+jdup+alyhqo9fAOZ2xkWTNNUtTrHubzpRlGh53kmTwjneWcBLLhEvSJr7enmJtgm+gqkbUrw=; ys=vbch.2-35-0; yashr=6774375601676365280; PHPSESSID=b958908d589c4751a619b529e1c1e461; SL_G_WPT_TO=ru; SL_GWPT_Show_Hide_tmp=1; SL_wptGlobTipTmp=1; is_gdpr=0; is_gdpr_b=CKWxOxDNpgE=; _yasc=IB2IJ3hkPKKs2lEqh2BZ4Rsk6eRqHOxSMhjHpjzKjXmp1e3xx4LRgi9QccB0ipw=; eda_web=%7B%22app%22%3A%7B%22lat%22%3A55.755245%2C%22lon%22%3A37.617779%2C%22deliveryTime%22%3Anull%2C%22xDeviceId%22%3A%22l6lo13q2-melqfb568r-57k4dzq0n59-l9q1wy8nfze%22%2C%22appBannerShown%22%3Afalse%2C%22isAdult%22%3Anull%2C%22yandexPlusCashbackOptInChecked%22%3Afalse%2C%22testRunId%22%3Anull%2C%22initialPromocode%22%3Anull%2C%22themeVariantKey%22%3A%22light%22%2C%22lang%22%3A%22ru%22%2C%22translateMenu%22%3Afalse%7D%7D',
    'origin': 'https://eda.yandex.ru',
    'referer': 'https://eda.yandex.ru/retail/paterocka/catalog/3442?placeSlug=pyaterochka_qnolu',
    'sec-ch-ua': '"Not_A Brand";v="99", "Google Chrome";v="109", "Chromium";v="109"',
    'sec-ch-ua-mobile': '?1',
    'sec-ch-ua-platform': '"Android"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
    'user-agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Mobile Safari/537.36',
    'x-app-version': '15.79.0',
    'x-device-id': 'l6lo13q2-melqfb568r-57k4dzq0n59-l9q1wy8nfze',
    'x-platform': 'mobile_web',
}


def start_search(address='qnolu', category=3374, font_size=15, root=""):
    json_data = {
        'slug': f'pyaterochka_{address}',
        'maxDepth': 200,
        'category_uid': category,
    }
    response = requests.post('https://eda.yandex.ru/api/v2/menu/goods', cookies=cookies, headers=headers,
                             json=json_data)

    subcategories, count = json.loads(response.text)['payload'].values()
    name_of_all_category = subcategories[0]["name"]
    create_excel(name_of_all_category, count,
                 list(sorted(subcategories[1:], key=lambda x: len(x["items"]), reverse=True)), font_size, root)


def create_excel(name_of_all_category, count, subcategories, font_size, root):
    workbook = xlsxwriter.Workbook(f'{name_of_all_category}-{datetime.today()}.xlsx')


    cell_format_level_1 = workbook.add_format()
    cell_format_level_1.set_font_size(15)
    cell_format_level_1.bold = 1

    cell_format_promo = workbook.add_format()
    cell_format_promo.set_bg_color('red')
    cell_format_promo.set_font_size(15)

    cell_format_level_2 = workbook.add_format()
    cell_format_level_2.set_font_size(15)

    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0, name_of_all_category, cell_format_level_1)
    i = 1
    for category in subcategories:
        category_name, goods = category["name"], category["items"]
        worksheet.write(i, 0, category_name, cell_format_level_1)
        i = i + 1
        for good in goods:
            worksheet.set_row(i, 30, None, {'level': 2, 'hidden': True})
            s = good['name'].split(' ')
            name, count = ' '.join(s[:-1]), s[-1]
            worksheet.write(i, 0, name, cell_format_level_2)
            if good.get('promoPrice'):
                worksheet.write(i, 1, f'Промо!!!!{good["promoPrice"]} ', cell_format_promo)
            else:
                worksheet.write(i, 1, good['decimalPrice'], cell_format_level_2)

            worksheet.write(i, 2, good["weight"], cell_format_level_2)
            worksheet.write(i, 3, count, cell_format_level_2)
            i = i + 1
    worksheet.set_column('A:A', 40)
    worksheet.set_column('B:B', 10)
    workbook.close()
