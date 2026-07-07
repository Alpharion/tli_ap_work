"""
coords.py — Comprehensive country → (lat, lng) lookup table (~200 countries + aliases).
Coordinates are approximate geographic centroids suitable for map dot placement.
"""

COUNTRY_COORDS = {
    # ── North America ─────────────────────────────────────────────────────────
    "united states": (37.09, -95.71), "usa": (37.09, -95.71), "us": (37.09, -95.71),
    "united states of america": (37.09, -95.71), "the united states": (37.09, -95.71),
    "u.s.a.": (37.09, -95.71), "u.s.": (37.09, -95.71), "america": (37.09, -95.71),
    "canada": (56.13, -106.34),
    "mexico": (23.63, -102.55),
    "cuba": (21.52, -77.78),
    "dominican republic": (18.73, -70.16),
    "puerto rico": (18.22, -66.59),
    "jamaica": (18.11, -77.29),
    "haiti": (18.97, -72.29),
    "trinidad and tobago": (10.69, -61.22),
    "trinidad": (10.69, -61.22),
    "bahamas": (25.03, -77.40),
    "barbados": (13.19, -59.54),
    "panama": (8.54, -80.78),
    "costa rica": (9.74, -83.75),
    "guatemala": (15.78, -90.23),
    "honduras": (15.20, -86.24),
    "el salvador": (13.79, -88.89),
    "nicaragua": (12.86, -85.21),
    "belize": (17.19, -88.49),

    # ── South America ────────────────────────────────────────────────────────
    "brazil": (-14.23, -51.92),
    "argentina": (-38.41, -63.61),
    "colombia": (4.57, -74.29),
    "chile": (-35.67, -71.54),
    "peru": (-9.19, -75.01),
    "venezuela": (6.42, -66.58),
    "ecuador": (-1.83, -78.18),
    "bolivia": (-16.29, -63.59),
    "paraguay": (-23.44, -58.44),
    "uruguay": (-32.52, -55.77),
    "guyana": (4.86, -58.93),
    "suriname": (3.92, -56.02),

    # ── Western Europe ───────────────────────────────────────────────────────
    "germany": (51.16, 10.45), "deutschland": (51.16, 10.45),
    "france": (46.22, 2.21),
    "united kingdom": (55.37, -3.43), "uk": (55.37, -3.43),
    "great britain": (55.37, -3.43), "britain": (55.37, -3.43),
    "england": (52.86, -1.46), "scotland": (56.49, -4.20),
    "wales": (52.13, -3.78), "northern ireland": (54.78, -6.49),
    "ireland": (53.41, -8.24), "republic of ireland": (53.41, -8.24),
    "netherlands": (52.13, 5.29), "the netherlands": (52.13, 5.29), "holland": (52.13, 5.29),
    "belgium": (50.50, 4.46),
    "luxembourg": (49.81, 6.13),
    "switzerland": (46.81, 8.22),
    "austria": (47.51, 14.55),
    "sweden": (60.12, 18.64),
    "norway": (60.47, 8.46),
    "denmark": (56.26, 9.50),
    "finland": (61.92, 25.74),
    "iceland": (64.96, -19.02),
    "spain": (40.46, -3.74),
    "portugal": (39.39, -8.22),
    "italy": (41.87, 12.56),
    "greece": (39.07, 21.82),
    "malta": (35.94, 14.37),
    "cyprus": (35.13, 33.43),
    "monaco": (43.74, 7.42),
    "liechtenstein": (47.14, 9.55),
    "andorra": (42.55, 1.57),
    "san marino": (43.94, 12.46),
    "vatican": (41.90, 12.45),

    # ── Central / Eastern Europe ─────────────────────────────────────────────
    "poland": (51.91, 19.14),
    "czech republic": (49.81, 15.47), "czechia": (49.81, 15.47), "czech": (49.81, 15.47),
    "slovakia": (48.66, 19.70), "slovak republic": (48.66, 19.70),
    "hungary": (47.16, 19.50),
    "romania": (45.94, 24.96),
    "bulgaria": (42.73, 25.49),
    "slovenia": (46.15, 14.99),
    "croatia": (45.10, 15.20),
    "serbia": (44.01, 21.00),
    "bosnia and herzegovina": (43.92, 17.68), "bosnia": (43.92, 17.68),
    "north macedonia": (41.60, 21.74), "macedonia": (41.60, 21.74),
    "albania": (41.15, 20.17),
    "montenegro": (42.71, 19.37),
    "kosovo": (42.60, 20.90),
    "ukraine": (48.37, 31.16),
    "moldova": (47.41, 28.37),
    "belarus": (53.70, 27.95),
    "lithuania": (55.16, 23.88),
    "latvia": (56.87, 24.60),
    "estonia": (58.59, 25.01),

    # ── Russia & CIS ─────────────────────────────────────────────────────────
    "russia": (61.52, 105.31), "russian federation": (61.52, 105.31),
    "kazakhstan": (48.01, 66.92),
    "uzbekistan": (41.37, 64.58),
    "turkmenistan": (38.97, 59.56),
    "kyrgyzstan": (41.20, 74.77),
    "tajikistan": (38.86, 71.28),
    "azerbaijan": (40.14, 47.57),
    "georgia": (42.31, 43.36),
    "armenia": (40.07, 45.04),

    # ── Middle East ──────────────────────────────────────────────────────────
    "turkey": (38.96, 35.24), "türkiye": (38.96, 35.24),
    "israel": (31.04, 34.85),
    "saudi arabia": (23.88, 45.07),
    "uae": (23.42, 53.84), "united arab emirates": (23.42, 53.84),
    "qatar": (25.35, 51.18),
    "kuwait": (29.31, 47.48),
    "bahrain": (26.00, 50.55),
    "oman": (21.51, 55.92),
    "jordan": (30.58, 36.23),
    "lebanon": (33.85, 35.86),
    "syria": (34.80, 38.99), "syrian arab republic": (34.80, 38.99),
    "iraq": (33.22, 43.68),
    "iran": (32.42, 53.68), "iran, islamic republic of": (32.42, 53.68),
    "yemen": (15.55, 48.52),
    "palestine": (31.95, 35.23), "west bank": (31.95, 35.23),

    # ── South Asia ───────────────────────────────────────────────────────────
    "india": (20.59, 78.96),
    "pakistan": (30.37, 69.34),
    "bangladesh": (23.68, 90.35),
    "sri lanka": (7.87, 80.77),
    "nepal": (28.39, 84.12),
    "bhutan": (27.51, 90.43),
    "maldives": (3.20, 73.22),
    "afghanistan": (33.93, 67.71),

    # ── East Asia ────────────────────────────────────────────────────────────
    "china": (35.86, 104.19), "prc": (35.86, 104.19), "people's republic of china": (35.86, 104.19),
    "japan": (36.20, 138.25),
    "south korea": (35.90, 127.76), "korea": (35.90, 127.76),
    "republic of korea": (35.90, 127.76), "korea, republic of": (35.90, 127.76),
    "korea, south": (35.90, 127.76),
    "north korea": (40.33, 127.51), "dprk": (40.33, 127.51),
    "taiwan": (23.69, 120.96), "taiwan, province of china": (23.69, 120.96),
    "hong kong": (22.32, 114.17), "hong kong sar": (22.32, 114.17),
    "macau": (22.16, 113.54), "macao": (22.16, 113.54),
    "mongolia": (46.86, 103.84),

    # ── Southeast Asia ───────────────────────────────────────────────────────
    "singapore": (1.35, 103.81),
    "malaysia": (4.21, 108.96),
    "thailand": (15.87, 100.99),
    "vietnam": (14.05, 108.27), "viet nam": (14.05, 108.27),
    "indonesia": (-0.78, 113.92),
    "philippines": (12.87, 121.77),
    "cambodia": (12.56, 104.99),
    "myanmar": (19.15, 96.63), "burma": (19.15, 96.63),
    "laos": (19.85, 102.50), "lao pdr": (19.85, 102.50), "lao people's democratic republic": (19.85, 102.50),
    "brunei": (4.53, 114.73),
    "timor-leste": (-8.87, 125.73), "east timor": (-8.87, 125.73),

    # ── Oceania ──────────────────────────────────────────────────────────────
    "australia": (-25.27, 133.77),
    "new zealand": (-40.90, 174.88),
    "papua new guinea": (-6.31, 143.96),
    "fiji": (-17.71, 178.06),

    # ── Africa ───────────────────────────────────────────────────────────────
    "south africa": (-30.55, 22.93),
    "nigeria": (9.08, 8.67),
    "egypt": (26.82, 30.80),
    "ethiopia": (9.14, 40.49),
    "kenya": (-0.02, 37.91),
    "ghana": (7.94, -1.02),
    "tanzania": (-6.37, 34.89), "united republic of tanzania": (-6.37, 34.89),
    "uganda": (1.37, 32.29),
    "mozambique": (-18.66, 35.53),
    "madagascar": (-18.77, 46.87),
    "cameroon": (7.37, 12.35),
    "angola": (-11.20, 17.87),
    "zimbabwe": (-19.01, 29.15),
    "zambia": (-13.13, 27.85),
    "senegal": (14.50, -14.45),
    "mali": (17.57, -3.99),
    "burkina faso": (12.36, -1.56),
    "guinea": (11.80, -15.18),
    "benin": (9.31, 2.32),
    "niger": (17.61, 8.08),
    "chad": (15.45, 18.73),
    "rwanda": (-1.94, 29.87),
    "burundi": (-3.37, 29.92),
    "somalia": (5.15, 46.20),
    "sudan": (12.86, 30.22),
    "south sudan": (6.87, 31.31),
    "democratic republic of congo": (-4.04, 21.76),
    "dr congo": (-4.04, 21.76), "drc": (-4.04, 21.76), "congo, democratic republic of the": (-4.04, 21.76),
    "congo": (-0.23, 15.83), "republic of congo": (-0.23, 15.83),
    "botswana": (-22.33, 24.68),
    "namibia": (-22.96, 18.49),
    "morocco": (31.79, -7.09),
    "algeria": (28.03, 1.66),
    "tunisia": (33.88, 9.54),
    "libya": (26.33, 17.23),
    "ivory coast": (7.54, -5.55), "côte d'ivoire": (7.54, -5.55), "cote d'ivoire": (7.54, -5.55),
    "liberia": (6.43, -9.43),
    "sierra leone": (8.46, -11.77),
    "eritrea": (15.18, 39.78),
    "djibouti": (11.83, 42.59),
    "gabon": (-0.80, 11.61),
    "equatorial guinea": (1.65, 10.27),
    "mauritius": (-20.35, 57.55),
    "seychelles": (-4.68, 55.49),
    "cape verde": (16.00, -24.01),
    "eswatini": (-26.52, 31.47), "swaziland": (-26.52, 31.47),
    "lesotho": (-29.61, 28.23),
    "malawi": (-13.25, 34.30),
}


def get_coords(country: str):
    if not country:
        return None
    key = country.strip().lower()
    # Direct lookup
    result = COUNTRY_COORDS.get(key)
    if result:
        return result
    # Strip trailing parenthetical or city info: "Shanghai, China" → "china"
    if "," in key:
        result = COUNTRY_COORDS.get(key.split(",")[-1].strip())
        if result:
            return result
        # Also try "city, country" → first part
        result = COUNTRY_COORDS.get(key.split(",")[0].strip())
        if result:
            return result
    # Strip common prefixes like "the "
    if key.startswith("the "):
        result = COUNTRY_COORDS.get(key[4:])
        if result:
            return result
    return None
