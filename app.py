from flask import Flask, request, send_file, render_template_string
from flask import Flask, render_template
from datetime import datetime, timedelta
import pytz
import pandas as pd
from weasyprint import HTML
import io
import swisseph as swe
from astral import LocationInfo
from astral.sun import sun
import openpyxl
from openpyxl.styles import PatternFill


###################################
from weasyprint import HTML
import io
###################################

app = Flask(__name__)

# ==============================================================================
# --- Configuration & Constants ---
# ==============================================================================
TZ = pytz.timezone("Asia/Kolkata")
LAT, LON, ALT = 26.17, 73.02, 0  # Jodhpur, Rajasthan coordinates
swe.set_ephe_path('./swefiles')
swe.set_sid_mode(swe.SIDM_LAHIRI)
swe.set_topo(LON, LAT, ALT)

NAKSHATRAS_INFO = [
    ("Ashwini", "Ketu"), ("Bharani", "Venus"), ("Krittika", "Sun"),
    ("Rohini", "Moon"), ("Mrigashirsha", "Mars"), ("Ardra", "Rahu"),
    ("Punarvasu", "Jupiter"), ("Pushya", "Saturn"), ("Ashlesha", "Mercury"),
    ("Magha", "Ketu"), ("Purva Phalguni", "Venus"), ("Uttara Phalguni", "Sun"),
    ("Hasta", "Moon"), ("Chitra", "Mars"), ("Swati", "Rahu"),
    ("Vishakha", "Jupiter"), ("Anuradha", "Saturn"), ("Jyeshtha", "Mercury"),
    ("Mula", "Ketu"), ("Purva Ashadha", "Venus"), ("Uttara Ashadha", "Sun"),
    ("Shravana", "Moon"), ("Dhanishta", "Mars"), ("Shatabhisha", "Rahu"),
    ("Purva Bhadrapada", "Jupiter"), ("Uttara Bhadrapada", "Saturn"), ("Revati", "Mercury")
]
NAKSHATRAS = [nak[0] for nak in NAKSHATRAS_INFO]
NAK_LORDS = [nak[1] for nak in NAKSHATRAS_INFO]

PLANET_IDS = {
    "Sun": swe.SUN, "Moon": swe.MOON, "Mars": swe.MARS, "Mercury": swe.MERCURY,
    "Jupiter": swe.JUPITER, "Venus": swe.VENUS, "Saturn": swe.SATURN,
    "Rahu": swe.MEAN_NODE, "Ketu": swe.TRUE_NODE
}
PLANET_ORDER = ["Sun", "Moon", "Mars", "Mercury", "Jupiter", "Venus", "Saturn", "Rahu", "Ketu"]

VEDHA_MAPPING = {
    "Ashwini": {"left": "Rohini", "front": "Purva Phalguni", "right": "Jyeshtha"},
    "Bharani": {"left": "Krittika", "front": "Magha", "right": "Anuradha"},
    "Krittika": {"left": "Vishakha", "front": "Shravana", "right": "Bharani"},
    "Rohini": {"left": "Swati", "front": "Abhijit", "right": "Ashwini"},
    "Mrigashirsha": {"left": "Chitra", "front": "Uttara Ashadha", "right": "Revati"},
    "Ardra": {"left": "Hasta", "front": "Purva Ashadha", "right": "Uttara Bhadrapada"},
    "Punarvasu": {"left": "Uttara Phalguni", "front": "Mula", "right": "Purva Bhadrapada"},
    "Pushya": {"left": "Purva Phalguni", "front": "Jyeshtha", "right": "Shatabhisha"},
    "Ashlesha": {"left": "Magha", "front": "Anuradha", "right": "Dhanishta"},
    "Magha": {"left": "Shravana", "front": "Bharani", "right": "Ashlesha"},
    "Purva Phalguni": {"left": "Abhijit", "front": "Ashwini", "right": "Pushya"},
    "Uttara Phalguni": {"left": "Uttara Ashadha", "front": "Revati", "right": "Punarvasu"},
    "Hasta": {"left": "Purva Ashadha", "front": "Uttara Bhadrapada", "right": "Ardra"},
    "Chitra": {"left": "Mula", "front": "Purva Bhadrapada", "right": "Mrigashirsha"},
    "Swati": {"left": "Jyeshtha", "front": "Shatabhisha", "right": "Rohini"},
    "Vishakha": {"left": "Anuradha", "front": "Krittika", "right": "Krittika"},
    "Anuradha": {"left": "Bharani", "front": "Ashlesha", "right": "Vishakha"},
    "Jyeshtha": {"left": "Ashwini", "front": "Pushya", "right": "Swati"},
    "Mula": {"left": "Revati", "front": "Punarvasu", "right": "Chitra"},
    "Purva Ashadha": {"left": "Uttara Bhadrapada", "front": "Ardra", "right": "Hasta"},
    "Uttara Ashadha": {"left": "Purva Bhadrapada", "front": "Mrigashirsha", "right": "Uttara Phalguni"},
    "Shravana": {"left": "Dhanishta", "front": "Krittika", "right": "Magha"},
    "Dhanishta": {"left": "Ashlesha", "front": "Vishakha", "right": "Shravana"},
    "Shatabhisha": {"left": "Pushya", "front": "Swati", "right": "Abhijit"},
    "Purva Bhadrapada": {"left": "Punarvasu", "front": "Chitra", "right": "Uttara Bhadrapada"},
    "Uttara Bhadrapada": {"left": "Ardra", "front": "Hasta", "right": "Purva Ashadha"},
    "Revati": {"left": "Mrigashirsha", "front": "Uttara Phalguni", "right": "Mula"}
}

SPEED_THRESHOLDS = {
    "Mercury": 59 / 60, "Jupiter": 5 / 60, "Venus": 59 / 60,
    "Mars": 31 / 60, "Saturn": 2 / 60
}

ZODIAC_SIGNS = [
    "Aries", "Taurus", "Gemini", "Cancer", "Leo", "Virgo",
    "Libra", "Scorpio", "Sagittarius", "Capricorn", "Aquarius", "Pisces"
]
RASHI_LORDS = {
    "Aries": "Mars", "Taurus": "Venus", "Gemini": "Mercury", "Cancer": "Moon",
    "Leo": "Sun", "Virgo": "Mercury", "Libra": "Venus", "Scorpio": "Mars",
    "Sagittarius": "Jupiter", "Capricorn": "Saturn", "Aquarius": "Saturn", "Pisces": "Jupiter"
}
RASHI_LORDS_INDEX = {
    0: "Mars", 1: "Venus", 2: "Mercury", 3: "Moon", 4: "Sun", 5: "Mercury",
    6: "Venus", 7: "Mars", 8: "Jupiter", 9: "Saturn", 10: "Saturn", 11: "Jupiter"
}
PERMANENT_FRIENDS = {
    "Sun": ["Moon", "Mars", "Jupiter"], "Moon": ["Sun", "Mercury"],
    "Mars": ["Sun", "Moon", "Jupiter", "Ketu"], "Mercury": ["Sun", "Venus"],
    "Jupiter": ["Sun", "Moon", "Mars"], "Venus": ["Mercury", "Saturn", "Rahu", "Ketu"],
    "Saturn": ["Mercury", "Venus", "Rahu"], "Rahu": ["Venus", "Saturn"],
    "Ketu": ["Mars", "Venus"]
}
PERMANENT_NEUTRALS = {
    "Sun": ["Mercury"], "Moon": ["Mars", "Jupiter", "Venus", "Saturn"],
    "Mars": ["Venus", "Saturn"], "Mercury": ["Mars", "Jupiter", "Saturn", "Rahu", "Ketu"],
    "Jupiter": ["Saturn", "Rahu", "Ketu"], "Venus": ["Mars", "Jupiter"],
    "Saturn": ["Jupiter"], "Rahu": ["Mercury", "Jupiter"], "Ketu": ["Mercury", "Jupiter"]
}
PLANETARY_SEQUENCE_HORA = ['Sun', 'Venus', 'Mercury', 'Moon', 'Saturn', 'Jupiter', 'Mars']
WEEKDAY_RULERS_HORA = ['Sun', 'Moon', 'Mars', 'Mercury', 'Jupiter', 'Venus', 'Saturn']

UPSIDE_BHAVS = {
    "Mercury": [1, 3, 4, 6, 7, 9, 12], "Venus": [6, 7, 9], "Sun": [1, 9, 12, 8],
    "Jupiter": [8, 12], "Mars": [1, 8, 9, 12], "Saturn": [8, 9, 12], "Moon": [1, 3, 6, 8, 12]
}
DOWNSIDE_BHAVS = {
    "Mercury": [2, 5, 8, 11, 10], "Venus": [8, 10, 11, 12], "Sun": [3, 6, 10, 11],
    "Jupiter": [9, 10, 11], "Mars": [3, 6, 10, 11], "Saturn": [3, 6, 10, 11],
    "Moon": [2, 4, 5, 7, 9, 10, 11]
}

ET1_PAIRS = [
    ("Gemini","Cancer"), ("Taurus","Leo"), ("Aries","Virgo"), ("Pisces","Libra"),
    ("Aquarius","Scorpio"), ("Capricorn","Sagittarius"), ("Cancer","Gemini"),
    ("Leo","Taurus"), ("Virgo","Aries"), ("Libra","Pisces"), ("Scorpio","Aquarius"),
    ("Sagittarius","Capricorn")
]
ET2_PAIRS = [
    ("Aries","Pisces"), ("Taurus","Aquarius"), ("Gemini","Capricorn"), ("Cancer","Sagittarius"),
    ("Leo","Scorpio"), ("Virgo","Libra"), ("Pisces","Aries"), ("Aquarius","Taurus"),
    ("Capricorn","Gemini"), ("Sagittarius","Cancer"), ("Scorpio","Leo"), ("Libra","Virgo")
]
COLOR_MAP = {
    "Upside": "#C6EFCE", "Downside": "#FFC7CE", "Good Downside": "#FFEB9C",
    "Rajyog + Downside": "#BDD7EE"
}
# ==============================================================================
# --- Core Calculation Functions ---
# ==============================================================================
def get_julian_day(dt_local):
    dt_utc = dt_local.astimezone(pytz.utc)
    return swe.julday(dt_utc.year, dt_utc.month, dt_utc.day,
                      dt_utc.hour + dt_utc.minute / 60.0 + dt_utc.second / 3600.0)

def get_rashi(degree):
    return ZODIAC_SIGNS[int(degree // 30)]

def get_nakshatra_info(degree):
    deg = degree % 360
    nak_degree_span = 360/27
    nak_index = int(deg / nak_degree_span)
    nak_name, nak_lord = "Unknown", "Unknown"
    if 0 <= nak_index < len(NAKSHATRAS_INFO):
        nak_name, nak_lord = NAKSHATRAS_INFO[nak_index]
    pada_degree_span = nak_degree_span / 4
    pada = int((deg % nak_degree_span) / pada_degree_span) + 1
    return nak_name, nak_lord, pada

def get_raw_longitude_and_speed(jd, planet_id):
    try:
        result = swe.calc_ut(jd, planet_id, swe.FLG_SWIEPH | swe.FLG_SPEED)
        return result[0][0], result[0][3]
    except Exception:
        return None, None

def get_planet_data(jd, ayanamsa):
    planet_data = {}
    rahu_raw_lon, rahu_speed = get_raw_longitude_and_speed(jd, PLANET_IDS['Rahu'])
    if rahu_raw_lon is not None:
        sidereal_lon_rahu = (rahu_raw_lon - ayanamsa) % 360
        nak_rahu, _, _ = get_nakshatra_info(sidereal_lon_rahu)
        planet_data['Rahu'] = { 'longitude': sidereal_lon_rahu, 'nakshatra': nak_rahu, 'speed': rahu_speed, 'retro': True }
        sidereal_lon_ketu = (rahu_raw_lon + 180 - ayanamsa) % 360
        nak_ketu, _, _ = get_nakshatra_info(sidereal_lon_ketu)
        planet_data['Ketu'] = { 'longitude': sidereal_lon_ketu, 'nakshatra': nak_ketu, 'speed': rahu_speed, 'retro': True }
    
    for p_name, p_id in PLANET_IDS.items():
        if p_name == "Rahu": continue
        raw_lon, speed = get_raw_longitude_and_speed(jd, p_id)
        if raw_lon is not None:
            sidereal_lon = (raw_lon - ayanamsa) % 360
            nak_name, _, _ = get_nakshatra_info(sidereal_lon)
            is_retro = speed < 0
            if p_name in ["Sun", "Moon"]: is_retro = False
            planet_data[p_name] = { 'longitude': sidereal_lon, 'nakshatra': nak_name, 'speed': speed, 'retro': is_retro }
    
    return planet_data
def get_planet_vedha_type_and_nakshatra(planet_name, nakshatra_name, is_retrograde, speed_in_degs_per_day):
    vedha_info = VEDHA_MAPPING.get(nakshatra_name)
    if not vedha_info: return None, "N/A"
    if planet_name in ["Sun", "Moon", "Rahu", "Ketu"]:
        return {"left": vedha_info.get('left'), "front": vedha_info.get('front'), "right": vedha_info.get('right')}, "All"
    elif planet_name in ["Mercury", "Venus", "Mars", "Jupiter", "Saturn"]:
        if is_retrograde: return vedha_info.get('right'), "Right"
        threshold = SPEED_THRESHOLDS.get(planet_name)
        if threshold is None or speed_in_degs_per_day <= threshold: return vedha_info.get('front'), "Front"
        return vedha_info.get('left'), "Left"
    return None, "N/A"

def calculate_asc_vedha_relationships(asc_nakshatra, planet_data):
    vedha_relationships = []
    for planet_name in PLANET_ORDER:
        if planet_name in planet_data:
            data = planet_data[planet_name]
            vedha_target, vedha_type = get_planet_vedha_type_and_nakshatra(
                planet_name, data['nakshatra'], data['retro'], data['speed'])
            if vedha_target:
                if vedha_type == "All":
                    for v_type, v_nak in vedha_target.items():
                        if v_nak == asc_nakshatra:
                            vedha_relationships.append({'planet': planet_name, 'vedha_type': v_type})
                elif vedha_target == asc_nakshatra:
                    vedha_relationships.append({'planet': planet_name, 'vedha_type': vedha_type})
    return vedha_relationships
def permanent_friendship_matrix():
    pf = {}
    for p1 in PLANET_ORDER:
        pf[p1] = {}
        for p2 in PLANET_ORDER:
            if p1 == p2: pf[p1][p2] = ''
            elif p2 in PERMANENT_FRIENDS.get(p1, []): pf[p1][p2] = "Friend"
            elif p2 in PERMANENT_NEUTRALS.get(p1, []): pf[p1][p2] = "Neutral"
            else: pf[p1][p2] = "Enemy"
    return pf
def get_hora_lord_at_datetime(dt_ist, latitude, longitude, timezone_str='Asia/Kolkata'):
    loc = LocationInfo(name="Custom", region="India", timezone=timezone_str, latitude=latitude, longitude=longitude)
    try:
        s_today = sun(loc.observer, date=dt_ist.date(), tzinfo=TZ)
        astrological_date = dt_ist.date() if dt_ist >= s_today['sunrise'] else dt_ist.date() - timedelta(days=1)
        s = sun(loc.observer, date=astrological_date, tzinfo=TZ)
        next_s = sun(loc.observer, date=astrological_date + timedelta(days=1), tzinfo=TZ)
        sunrise, sunset, next_sunrise = s['sunrise'], s['sunset'], next_s['sunrise']
    except Exception: return "N/A"
    
    day_duration_seconds = (sunset - sunrise).total_seconds()
    night_duration_seconds = (next_sunrise - sunset).total_seconds()
    if day_duration_seconds <= 0 or night_duration_seconds <= 0: return "N/A"
    weekday = astrological_date.weekday()
    day_ruler_index = (weekday + 1) % 7
    day_ruler = WEEKDAY_RULERS_HORA[day_ruler_index]
    start_index_planetary_sequence = PLANETARY_SEQUENCE_HORA.index(day_ruler)
    if sunrise <= dt_ist < sunset:
        elapsed_seconds = (dt_ist - sunrise).total_seconds()
        hora_duration_seconds = day_duration_seconds / 12
        hora_number = int(elapsed_seconds // hora_duration_seconds)
        hora_planet_index = (start_index_planetary_sequence + hora_number) % 7
        return PLANETARY_SEQUENCE_HORA[hora_planet_index]
    elif sunset <= dt_ist < next_sunrise:
        elapsed_seconds = (dt_ist - sunset).total_seconds()
        hora_duration_seconds = night_duration_seconds / 12
        hora_number = int(elapsed_seconds // hora_duration_seconds)
        hora_planet_index = (start_index_planetary_sequence + 12 + hora_number) % 7
        return PLANETARY_SEQUENCE_HORA[hora_planet_index]
    return "N/A"

def calculate_hora_bhav_info(hora_lord, asc_rashi, planet_data):
    if hora_lord == "N/A": return "N/A", "N/A", "N/A"
    hora_lord_data = planet_data.get(hora_lord)
    if not hora_lord_data: return "N/A", "N/A", "N/A"
    hora_lord_sign = get_rashi(hora_lord_data['longitude'])
    asc_sign_index = ZODIAC_SIGNS.index(asc_rashi)
    hora_lord_sign_index = ZODIAC_SIGNS.index(hora_lord_sign)
    hora_bhav_number = (hora_lord_sign_index - asc_sign_index + 12) % 12 + 1
    return hora_lord, hora_bhav_number, hora_lord_sign

def calculate_ud_signal_friendship(planet1, planet2, permanent_friendship_data):
    if planet1 == "N/A" or planet2 == "N/A": return "Neutral"
    friendship_status = permanent_friendship_data.get(planet1, {}).get(planet2, "Neutral")
    if friendship_status == "Friend": return "Downside"
    if friendship_status == "Enemy": return "Upside"
    return "Neutral"

def calculate_ud_signal_bhav(hora_lord, hora_bhav_number):
    if hora_lord == "N/A" or hora_bhav_number == "N/A": return "Neutral"
    if hora_lord in UPSIDE_BHAVS and hora_bhav_number in UPSIDE_BHAVS[hora_lord]: return "Upside"
    if hora_lord in DOWNSIDE_BHAVS and hora_bhav_number in DOWNSIDE_BHAVS[hora_lord]: return "Downside"
    return "Neutral"

def calculate_final_logic_of_ud(ud_hora_asc_nak_lord, ud_hora_vedha_planet_signals, all_signals):
    if ud_hora_asc_nak_lord != "Neutral": return ud_hora_asc_nak_lord
    if ud_hora_vedha_planet_signals:
        non_neutral = [s for s in ud_hora_vedha_planet_signals if s != "Neutral"]
        if non_neutral:
            ups = non_neutral.count("Upside")
            downs = non_neutral.count("Downside")
            if ups > downs: return "Upside"
            if downs > ups: return "Downside"
            return "Neutral"
    ups = all_signals.count("Upside")
    downs = all_signals.count("Downside")
    if ups > downs: return "Upside"
    if downs > ups: return "Downside"
    return "Neutral"

def norm_deg(x): return (x % 360.0 + 360) % 360.0
def get_sign_index_from_lon(lon): return int(norm_deg(lon) // 30)
def get_nakshatra_index_from_lon(lon): return int(norm_deg(lon) // (360.0/27.0))
def get_charan_from_lon(lon):
    nak_size = 360.0/27.0
    return int((norm_deg(lon) % nak_size) // (nak_size / 4.0)) + 1
def get_navamsa_long_from_lon(lon): return norm_deg(lon * 9.0)
def get_navamsa_sign_index(lon): return int(get_navamsa_long_from_lon(lon) // 30)
def calc_planet_lon(jd_ut, planet_id):
    pos, _ = swe.calc_ut(jd_ut, planet_id, swe.FLG_SIDEREAL)
    return norm_deg(pos[0])

def get_details_for_planet_at_jd(jd_ut, planet_id, planet_name):
    lon = calc_planet_lon(jd_ut, planet_id)
    rashi_idx = get_sign_index_from_lon(lon)
    nak_idx = get_nakshatra_index_from_lon(lon)
    nav_idx = get_navamsa_sign_index(lon)
    nav_lon = get_navamsa_long_from_lon(lon)
    nav_nak_idx = get_nakshatra_index_from_lon(nav_lon)
    return {
        "rashi": ZODIAC_SIGNS[rashi_idx], "rashi_lord": RASHI_LORDS_INDEX[rashi_idx],
        "nakshatra": NAKSHATRAS[nak_idx], "nakshatra_lord": NAK_LORDS[nak_idx],
        "charan": get_charan_from_lon(lon), "nav_rashi": ZODIAC_SIGNS[nav_idx],
        "nav_lord": RASHI_LORDS_INDEX[nav_idx], "nav_nakshatra": NAKSHATRAS[nav_nak_idx],
        "nav_nak_lord": NAK_LORDS[nav_nak_idx], "nav_charan": get_charan_from_lon(nav_lon)
    }

def get_planet_key(dt_ist, planet_name):
    jd_ut = get_julian_day(dt_ist)
    lon = calc_planet_lon(jd_ut, PLANET_IDS[planet_name])
    return (ZODIAC_SIGNS[get_sign_index_from_lon(lon)],
            NAKSHATRAS[get_nakshatra_index_from_lon(lon)],
            get_charan_from_lon(lon),
            ZODIAC_SIGNS[get_navamsa_sign_index(lon)])

def relation(central_planet, other_planet_name):
    if other_planet_name in PERMANENT_FRIENDS.get(central_planet, []): return "Friend"
    if other_planet_name in PERMANENT_NEUTRALS.get(central_planet, []): return "Neutral"
    return "Enemy"

def check_ET_pairs(center_sign, other_sign, et_pairs):
    return any((center_sign == a and other_sign == b) or (center_sign == b and other_sign == a) for a,b in et_pairs)

def part3_remark(central, et_planet_list, d1_lord, d9_lord):
    planets = {entry.split()[0] for entry in et_planet_list}
    if central == "Sun":
        if planets.intersection({"Saturn", "Rahu", "Ketu"}): return "Upside"
        if "Jupiter" in planets: return "Downside"
    elif central == "Moon":
        if planets.intersection({"Rahu", "Ketu"}): return "Upside"
        if "Saturn" in planets: return "Rajyog + Downside"
        if "Mercury" in planets: return "Good Downside" if d1_lord == "Mercury" or d9_lord == "Mercury" else "Downside"
        if "Jupiter" in planets: return "Downside"
    return ""

def get_full_details_row(dt_ist, central):
    jd_ut = get_julian_day(dt_ist)
    details = {pname: get_details_for_planet_at_jd(jd_ut, pid, pname) for pname, pid in PLANET_IDS.items()}
    cent = details[central]
    et1_d1, et2_d1, et1_d9, et2_d9 = [], [], [], []
    other_planets = [p for p in PLANET_ORDER if p != central]

    for op in other_planets:
        op_det = details.get(op)
        if not op_det:
            if op == "Ketu":
                ayanamsa = swe.get_ayanamsa_ut(jd_ut)
                rahu_lon_raw, _ = get_raw_longitude_and_speed(jd_ut, PLANET_IDS['Rahu'])
                ketu_lon = norm_deg((rahu_lon_raw - ayanamsa) + 180.0)
                op_det = {
                    "rashi": ZODIAC_SIGNS[get_sign_index_from_lon(ketu_lon)],
                    "nav_rashi": ZODIAC_SIGNS[get_navamsa_sign_index(ketu_lon)]
                }
            else:
                continue
        
        if check_ET_pairs(cent["rashi"], op_det["rashi"], ET1_PAIRS): et1_d1.append(f"{op} ({relation(central, op)})")
        if check_ET_pairs(cent["rashi"], op_det["rashi"], ET2_PAIRS): et2_d1.append(f"{op} ({relation(central, op)})")
        if check_ET_pairs(cent["nav_rashi"], op_det["nav_rashi"], ET1_PAIRS): et1_d9.append(f"{op} ({relation(central, op)})")
        if check_ET_pairs(cent["nav_rashi"], op_det["nav_rashi"], ET2_PAIRS): et2_d9.append(f"{op} ({relation(central, op)})")
    
    remark = part3_remark(central, et1_d1 + et2_d1 + et1_d9 + et2_d9, cent["rashi_lord"], cent["nav_lord"])
    return {
        "DateTime (IST)": dt_ist.strftime("%Y-%m-%d %H:%M:%S"),
        **cent,
        "D1 ET1": ", ".join(et1_d1), "D1 ET2": ", ".join(et2_d1),
        "D9 ET1": ", ".join(et1_d9), "D9 ET2": ", ".join(et2_d9),
        "Part3 Remark": remark
    }
# ==============================================================================
# --- Report Generation Functions ---
# ==============================================================================
def generate_ascendant_report(start_dt_str, end_dt_str):
    output_rows = []
    start_dt_full = TZ.localize(datetime.strptime(start_dt_str, "%Y-%m-%dT%H:%M"))
    end_dt_full = TZ.localize(datetime.strptime(end_dt_str, "%Y-%m-%dT%H:%M"))
    current_dt = start_dt_full
    step_minutes = 2
    
    previous_state = {"Nakshatra": None, "Pada": None, "Logic_of_U_D": None}
    permanent_friendship_data = permanent_friendship_matrix()

    while current_dt <= end_dt_full:
        jd = get_julian_day(current_dt)
        ayanamsa = swe.get_ayanamsa_ut(jd)
        
        tropical_asc_deg = swe.houses_ex(jd, LAT, LON, b'P')[0][0]
        asc_deg = (tropical_asc_deg - ayanamsa) % 360
        asc_rashi = get_rashi(asc_deg)
        asc_nak, asc_nak_lord, asc_nak_pad = get_nakshatra_info(asc_deg)
        asc_rashi_lord = RASHI_LORDS.get(asc_rashi, "Unknown")
        
        planet_data = get_planet_data(jd, ayanamsa)
        vedha_rels = calculate_asc_vedha_relationships(asc_nak, planet_data)
        
        hora_lord = get_hora_lord_at_datetime(current_dt, LAT, LON, TZ.zone)
        hora_lord_bhav, hora_bhav_num, _ = calculate_hora_bhav_info(hora_lord, asc_rashi, planet_data)

        ud_hora_rashi_lord = calculate_ud_signal_friendship(hora_lord, asc_rashi_lord, permanent_friendship_data)
        ud_hora_vedha_planet_signals = [calculate_ud_signal_friendship(hora_lord, r['planet'], permanent_friendship_data) for r in vedha_rels]
        ud_hora_nak_lord = calculate_ud_signal_friendship(hora_lord, asc_nak_lord, permanent_friendship_data)
        ud_asc_bhav = calculate_ud_signal_bhav(hora_lord, hora_bhav_num)
        
        logic_ud = ""
        is_override = False
        if hora_lord == "Moon" and any(r['planet'] == "Moon" for r in vedha_rels):
            logic_ud = "Downside"
            is_override = True
        elif hora_lord == "Sun" and any(r['planet'] in ["Saturn", "Rahu", "Ketu"] for r in vedha_rels):
            logic_ud = "Upside"
            is_override = True
        elif hora_lord == "Moon" and any(r['planet'] in ["Rahu", "Ketu"] for r in vedha_rels):
            logic_ud = "Upside"
            is_override = True
        if not is_override:
            all_signals = [ud_hora_rashi_lord, ud_hora_nak_lord, ud_asc_bhav] + ud_hora_vedha_planet_signals
            # logic_ud = calculate_final_logic_of_ud(ud_hora_asc_nak_lord, ud_hora_vedha_planet_signals, all_signals)
            logic_ud = calculate_final_logic_of_ud(ud_hora_nak_lord, ud_hora_vedha_planet_signals, all_signals)

        current_state = {
            "Date": current_dt.strftime("%Y-%m-%d"),
            "Time": current_dt.strftime("%H:%M"),
            "U/D Logic": logic_ud
        }
        
        # --- CORRECTED CHANGE DETECTION LOGIC ---
        # The key is to check if Nakshatra or Pada has changed, OR if the U/D Logic has changed.
        is_changed = (
            asc_nak != previous_state.get("Nakshatra") or
            asc_nak_pad != previous_state.get("Pada") or
            logic_ud != previous_state.get("Logic_of_U_D")
        )
        
        if is_changed:
            output_rows.append(current_state)
            previous_state["Nakshatra"] = asc_nak
            previous_state["Pada"] = asc_nak_pad
            previous_state["Logic_of_U_D"] = logic_ud

        current_dt += timedelta(minutes=step_minutes)

    df = pd.DataFrame(output_rows)
    return df

def generate_sun_moon_report(start_dt_str, end_dt_str):
    sun_rows = []
    moon_rows = []
    start_ist = TZ.localize(datetime.strptime(start_dt_str, "%Y-%m-%dT%H:%M"))
    end_ist = TZ.localize(datetime.strptime(end_dt_str, "%Y-%m-%dT%H:%M"))
    current_ist = start_ist
    step_minutes = 1
    
    prev_sun_key = get_planet_key(current_ist, "Sun")
    prev_moon_key = get_planet_key(current_ist, "Moon")
    sun_rows.append(get_full_details_row(current_ist, "Sun"))
    moon_rows.append(get_full_details_row(current_ist, "Moon"))

    while current_ist < end_ist:
        next_ist = min(current_ist + timedelta(minutes=step_minutes), end_ist)
        
        curr_sun_key = get_planet_key(next_ist, "Sun")
        if curr_sun_key != prev_sun_key:
            low, high = current_ist, next_ist
            while (high - low).total_seconds() > 1:
                mid = low + (high - low) / 2
                if get_planet_key(mid, "Sun") == prev_sun_key: low = mid
                else: high = mid
            sun_rows.append(get_full_details_row(high, "Sun"))
            prev_sun_key = curr_sun_key

        curr_moon_key = get_planet_key(next_ist, "Moon")
        if curr_moon_key != prev_moon_key:
            low, high = current_ist, next_ist
            while (high - low).total_seconds() > 1:
                mid = low + (high - low) / 2
                if get_planet_key(mid, "Moon") == prev_moon_key: low = mid
                else: high = mid
            moon_rows.append(get_full_details_row(high, "Moon"))
            prev_moon_key = curr_moon_key
        
        current_ist = next_ist

    # df_sun = pd.DataFrame(sun_rows)
    # df_sun['Time'] = df_sun['DateTime (IST)'].apply(lambda x: x.split(' ')[1])
    # df_sun = df_sun[['Time', 'Part3 Remark']]
    
    # df_moon = pd.DataFrame(moon_rows)
    # df_moon['Time'] = df_moon['DateTime (IST)'].apply(lambda x: x.split(' ')[1])
    # df_moon = df_moon[['Time', 'Part3 Remark']]
    df_sun = pd.DataFrame(sun_rows)
    df_sun = df_sun[['DateTime (IST)', 'Part3 Remark']].rename(columns={'DateTime (IST)': 'Date & Time'})
    
    df_moon = pd.DataFrame(moon_rows)
    df_moon = df_moon[['DateTime (IST)', 'Part3 Remark']].rename(columns={'DateTime (IST)': 'Date & Time'})

    return df_sun, df_moon

# def generate_pdf(dataframe, title, color_column=None, color_map=None):
#     if dataframe.empty:
#         return None
#     def style_rows(row):
#         style = ''
#         if color_column and color_column in row.index and color_map:
#             value = row[color_column]
#             if value and value in color_map:
#                 hex_color = color_map[value]
#                 style = f'background-color: {hex_color};'
#         return [style] * len(row)
#     styled_df = dataframe.style.apply(style_rows, axis=1) if color_column else dataframe.style
#     html_content = styled_df.hide(axis='index').to_html()
#     full_html = f"""
#     <!DOCTYPE html>
#     <html>
#     <head>
#         <meta charset="UTF-8">
#         <title>{title}</title>
#         <style>
#             @page {{ size: A4 portrait; margin: 1cm; }}
#             body {{ font-family: 'Helvetica', 'Arial', sans-serif; }}
#             h1 {{ text-align: center; color: #333; }}
#             table {{ border-collapse: collapse; width: 100%; font-size: 10pt; }}
#             th, td {{ border: 1px solid #cccccc; text-align: left; padding: 8px; }}
#             th {{ background-color: #f2f2f2; font-weight: bold; }}
#         </style>
#     </head>
#     <body>
#         <h1>{title}</h1>
#         {html_content}
#     </body>
#     </html>
#     """
#     buffer = io.BytesIO()
#     HTML(string=full_html).write_pdf(buffer)
#     buffer.seek(0)
#     return buffer


def generate_pdf(dataframe, title, color_column=None, color_map=None):
    if dataframe.empty:
        return None

    # Apply row-wise styling for background colors
    def style_rows(row):
        style = ''
        if color_column and color_column in row.index and color_map:
            value = row[color_column]
            if value and value in color_map:
                hex_color = color_map[value]
                style = f'background-color: {hex_color};'
        return [style] * len(row)

    styled_df = dataframe.style.apply(style_rows, axis=1) if color_column else dataframe.style
    html_content = styled_df.hide(axis='index').to_html()

    # Full HTML document
    full_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>{title}</title>
        <style>
            @page {{ size: A4 portrait; margin: 1cm; }}
            body {{ font-family: 'Helvetica', 'Arial', sans-serif; }}
            h1 {{ text-align: center; color: #333; }}
            table {{ border-collapse: collapse; width: 100%; font-size: 10pt; }}
            th, td {{ border: 1px solid #cccccc; text-align: left; padding: 8px; }}
            th {{ background-color: #f2f2f2; font-weight: bold; }}
        </style>
    </head>
    <body>
        <h1>{title}</h1>
        {html_content}
    </body>
    </html>
    """

    # Convert HTML to PDF using WeasyPrint
    pdf_buffer = io.BytesIO()
    HTML(string=full_html).write_pdf(pdf_buffer)
    pdf_buffer.seek(0)
    return pdf_buffer


@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate-report', methods=['POST'])
def generate_report():
    data = request.json
    start_dt_str = data.get('start_datetime')
    end_dt_str = data.get('end_datetime')
    report_type = data.get('report_type')
    
    if not start_dt_str or not end_dt_str or not report_type:
        return "Missing required parameters", 400

    # Convert the input strings to datetime objects for comparison
    start_dt_full = TZ.localize(datetime.strptime(start_dt_str, "%Y-%m-%dT%H:%M"))
    end_dt_full = TZ.localize(datetime.strptime(end_dt_str, "%Y-%m-%dT%H:%M"))

    # Set the allowed date range for manual restriction
    min_allowed_date = TZ.localize(datetime(2025, 1, 1))
    max_allowed_date = TZ.localize(datetime(2025, 10, 31, 23, 59, 59))

    # Check if the requested date range is outside the allowed period
    if not (min_allowed_date <= start_dt_full <= max_allowed_date and
            min_allowed_date <= end_dt_full <= max_allowed_date):
        return "Date range is outside the allowed period (Jan 1, 2025 - Oct 31, 2025).", 400

    if report_type == 'asc':
        df = generate_ascendant_report(start_dt_str, end_dt_str)
        pdf_buffer = generate_pdf(df, "Ascendant U/D Report", color_column="U/D Logic", color_map=COLOR_MAP)
        if pdf_buffer:
            return send_file(pdf_buffer, mimetype='application/pdf', as_attachment=True, download_name='ascendant_report.pdf')
        else:
            return "No data to generate report.", 404
    elif report_type == 'sun_moon':
        df_sun, df_moon = generate_sun_moon_report(start_dt_str, end_dt_str)
        
        # merged_df = pd.concat([
        #     pd.DataFrame([{'Time': 'SUN REPORT', 'Part3 Remark': ''}]), df_sun,
        #     pd.DataFrame([{'Time': '', 'Part3 Remark': ''}]),
        #     pd.DataFrame([{'Time': 'MOON REPORT', 'Part3 Remark': ''}]), df_moon
        # ], ignore_index=True)
        # merged_df = merged_df.rename(columns={'Part3 Remark': 'Remarks'})
        merged_df = pd.concat([
        pd.DataFrame([{'Date & Time': 'SUN REPORT', 'Part3 Remark': ''}]), df_sun,
        pd.DataFrame([{'Date & Time': '', 'Part3 Remark': ''}]),
        pd.DataFrame([{'Date & Time': 'MOON REPORT', 'Part3 Remark': ''}]), df_moon
         ], ignore_index=True)
        merged_df = merged_df.rename(columns={'Part3 Remark': 'Remarks'})

        pdf_buffer = generate_pdf(merged_df, "Sun & Moon Report", color_column="Remarks", color_map=COLOR_MAP)
        if pdf_buffer:
            return send_file(pdf_buffer, mimetype='application/pdf', as_attachment=True, download_name='sun_moon_report.pdf')
        else:
            return "No data to generate report.", 404
    else:
        return "Invalid report type", 400

if __name__ == '__main__':
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)

