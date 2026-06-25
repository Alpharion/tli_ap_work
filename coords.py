"""
coords.py — Static country → (lat, lng) lookup table.
"""

COUNTRY_COORDS = {
    "united states":(37.09,-95.71),"usa":(37.09,-95.71),"us":(37.09,-95.71),
    "china":(35.86,104.19),"prc":(35.86,104.19),
    "japan":(36.20,138.25),"germany":(51.16,10.45),
    "south korea":(35.90,127.76),"korea":(35.90,127.76),
    "taiwan":(23.69,120.96),"united kingdom":(55.37,-3.43),"uk":(55.37,-3.43),
    "france":(46.22,2.21),"india":(20.59,78.96),"canada":(56.13,-106.34),
    "australia":(-25.27,133.77),"brazil":(-14.23,-51.92),"mexico":(23.63,-102.55),
    "netherlands":(52.13,5.29),"sweden":(60.12,18.64),"switzerland":(46.81,8.22),
    "italy":(41.87,12.56),"spain":(40.46,-3.74),"singapore":(1.35,103.81),
    "malaysia":(4.21,108.96),"thailand":(15.87,100.99),"vietnam":(14.05,108.27),
    "indonesia":(-0.78,113.92),"philippines":(12.87,121.77),"russia":(61.52,105.31),
    "saudi arabia":(23.88,45.07),"uae":(23.42,53.84),"united arab emirates":(23.42,53.84),
    "south africa":(-30.55,22.93),"nigeria":(9.08,8.67),"egypt":(26.82,30.80),
    "turkey":(38.96,35.24),"poland":(51.91,19.14),"czech republic":(49.81,15.47),
    "hungary":(47.16,19.50),"finland":(61.92,25.74),"norway":(60.47,8.46),
    "denmark":(56.26,9.50),"belgium":(50.50,4.46),"austria":(47.51,14.55),
    "portugal":(39.39,-8.22),"israel":(31.04,34.85),"new zealand":(-40.90,174.88),
    "argentina":(-38.41,-63.61),"chile":(-35.67,-71.54),"colombia":(4.57,-74.29),
    "bangladesh":(23.68,90.35),"pakistan":(30.37,69.34),"sri lanka":(7.87,80.77),
    "cambodia":(12.56,104.99),"myanmar":(19.15,96.63),
}


def get_coords(country: str):
    if not country:
        return None
    key = country.strip().lower()
    return COUNTRY_COORDS.get(key) or COUNTRY_COORDS.get(key.split(",")[0].strip())
