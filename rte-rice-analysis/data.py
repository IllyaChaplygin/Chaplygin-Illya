# -*- coding: utf-8 -*-
"""Master dataset: RTE rice Ukraine — market SKUs, channels, own cost model."""

# ---- design tokens (from RTE_Rice_Market_Research..pptx) ----
INK      = "#1C2235"; INK2 = "#303A5D"; MUTED = "#717890"
MUTED2   = "#9FA9C4"; LINE = "#D5DBEA"; BG = "#F1F4FA"; BG2 = "#E2E8F3"; WHITE="#FFFFFF"
# validated categorical palette (six checks PASS, light mode)
C_AMBER="#C4820A"; C_TEAL="#1E9EA6"; C_PURPLE="#5D4B96"; C_PINK="#C94F7C"
C_GREEN="#37A169"; C_BRAND="#F9A50B"

# ---- model constants (from RICE_PRICING_Slelf_.xlsx) ----
FX          = 45.0      # грн/$
BONUS       = 0.25      # бонус мережі
MARGIN      = 0.35      # цільова маржа
MARKUP      = 1.40      # націнка мережі
CC_SHARE    = 1 - BONUS - MARGIN        # 0.40
SHELF_MULT  = MARKUP / CC_SHARE         # 3.5  -> полиця = СС_грн * 3.5
# landed-cost multiplier on FOB, derived from suppliers deck (20'FCL bulk, Bangkok)
K_LANDED_20 = 0.995418606770833 / 0.583      # 1.7075
K_LANDED_40 = 0.892315648148148 / 0.583      # 1.5306
FOB_TO_SHELF_20 = K_LANDED_20 * FX * SHELF_MULT   # 268.9
FOB_TO_SHELF_40 = K_LANDED_40 * FX * SHELF_MULT   # 241.0

# ---- market: packaging formats (84 SKU, 22-23.09.2026) ----
# name, positions, share%, brands, median_price_uah, lo, hi, median_weight_g
FORMATS = [
    ("Пауч · реторт",       19, 23, 5, 109, 45,  376, 250, C_AMBER),
    ("Чаша · стакан",       23, 27, 5, 229, 83,  931, 218, C_TEAL),
    ("Дойпак · сублімат",   30, 36, 11, 315, 55,  889, 110, C_PURPLE),
    ("Коробка з нагрівачем",12, 14, 3, 708, 380, 1190, 230, C_PINK),
]

# ---- market: brand matrix (23 brands) ----
# brand, country, tech, format, sku, lo, hi, median, w_lo, w_hi, uah_per_100g_lo, hi, channel_type
BRANDS = [
 ("Ben's Original","ЄС","Розігрів","Пауч",13, 45, 179, 90, 220,250, 18, 81,"Мережа + маркетплейс"),
 ("Ottogi","Корея","Розігрів","Чаша",10, 225, 356, 255, 217,320, 79,130,"Азійські фудшопи"),
 ("Henan","Китай","Окріп","Чаша",  8,  83, 156, 140, 144,174, 58,106,"Азійські фудшопи + МП"),
 ("Bibigo","Корея","Розігрів","Чаша",1, 135, 189, 162, 210,210, 64, 90,"Фудшопи + маркетплейс"),
 ("Clearspring","ЄС","Розігрів","Пауч",1, 252, 446, 349, 250,250,101,178,"Маркетплейс + спец."),
 ("Portion","Україна","Розігрів","Пауч",1, 109, 109, 109, 350,350, 31, 31,"Маркетплейс"),
 ("Маркел","Україна","Розігрів","Пауч",1,  94,  96,  95, 350,350, 27, 27,"Військторг"),
 ("Gallina Blanca","ЄС","Окріп","Чаша",2, 229, 229, 229,  84, 84,273,273,"Спец. магазин"),
 ("Qiaoshanmei","Китай","Окріп","Пакет",2, 501, 501, 501, 146,146,343,343,"Азійський фудшоп"),
 ("Haidilao","Китай","Саморозігрів","Коробка",10,380,1190,680,165,360,230,371,"Туризм + фудшопи"),
 ("Mo Xiao Xian","Китай","Саморозігрів","Чаша",2, 931, 931, 931, 275,275,339,339,"Азійський фудшоп"),
 ("Zihaiguo","Китай","Саморозігрів","Коробка",1, 735, 735, 735, 440,440,167,167,"Азійський фудшоп"),
 ("Rongcheng Haoji","Китай","Саморозігрів","Коробка",1,735,735,735,440,440,167,167,"Азійський фудшоп"),
 ("Travellunch","ЄС","Сублімація","Дойпак",7, 275, 889, 515, 125,250,220,454,"Туризм"),
 ("Trek'n Eat","ЄС","Сублімація","Дойпак",2, 269, 715, 492,   0,  0,  0,  0,"Туризм"),
 ("Mountain House","США","Сублімація","Дойпак",2, 699, 699, 699, 110,133,526,635,"Туризм"),
 ("Adventure Food","ЄС","Сублімація","Дойпак",1, 492, 492, 492, 146,146,337,337,"Туризм"),
 ("Adventure Menu","ЄС","Розігрів+Субл.","Пауч/Дойпак",5,315,483,378,110,400,79,439,"Туризм"),
 ("SubliMate","Україна","Сублімація","Дойпак",3, 290, 370, 310, 110,140,221,308,"Туризм"),
 ("James Cook","Україна","Сублімація","Дойпак",4,  55, 156, 133,  80, 90, 61,195,"Туризм"),
 ("Їжа в Похід","Україна","Сублімація","Дойпак",2, 130, 145, 138,  85, 85,153,171,"Власний магазин"),
 ("Харчі","Україна","Сублімація","Дойпак",4,  83, 321, 120,  85,100, 98,321,"Туризм"),
 ("!FEST","Україна","Сублімація","Дойпак",1, 133, 140, 137, 100,100,133,140,"Туризм"),
]

# ---- adjacent shelf category (консерва/плов), slide 17 ----
ADJACENT = [
 ("hapay! каша рисова зі свининою","hapay!",340, 63, 18.5,"Ашан"),
 ("L'appetit каша зі свининою","L'appetit",340,120, 35.3,"Ашан"),
 ("L'appetit каша з яловичиною","L'appetit",340,126, 37.1,"Ашан"),
 ("hapay! плов з качки та булгуру","hapay!",340,117, 34.4,"Ашан"),
 ("Food Fabrika плов з куркою","Food Fabrika",250,119, 47.6,"Восторг"),
 ("М'ясторія плов з куркою","М'ясторія",350,153, 43.7,"Novus·МегаМаркет·Космос"),
]

# ---- own portfolio: cost model (20'FCL bulk) ----
# group, sku, weight_g, cc_usd_20, cc_usd_40, supplier, fmt
OWN = [
 ("BSCM — пауч 240 г","Jasmine Rice",240,0.995418606770833,0.892315648148148,"BSCM","Пауч"),
 ("BSCM — пауч 240 г","Mexican Rice",240,1.13762102083333,1.0197892037037,"BSCM","Пауч"),
 ("BSCM — пауч 240 г","Japanese Style Rice",240,1.13762102083333,1.0197892037037,"BSCM","Пауч"),
 ("BSCM — пауч 240 г","Hot and Spicy Rice",240,1.27982343489583,1.14726275925926,"BSCM","Пауч"),
 ("BSCM — пауч 200 г","Jasmine Rice (Pathum)",200,0.952556220056099,0.846095788605277,"BSCM","Пауч"),
 ("BSCM — пауч 200 г","Mexican Rice",200,1.07379043215039,0.953780613255561,"BSCM","Пауч"),
 ("BSCM — пауч 200 г","Japanese Style Rice",200,1.07379043215039,0.953780613255561,"BSCM","Пауч"),
 ("BSCM — пауч 200 г","Hot and Spicy Rice",200,1.17770547108835,1.04608189152723,"BSCM","Пауч"),
 ("BSCM — пауч 150 г","Jasmine Rice (Pathum)",150,0.842849973738682,0.738866431789378,"BSCM","Пауч"),
 ("BSCM — пауч 150 г","Mexican Rice",150,0.981094929271472,0.860056042975836,"BSCM","Пауч"),
 ("BSCM — пауч 150 г","Japanese Style Rice",150,0.981094929271472,0.860056042975836,"BSCM","Пауч"),
 ("BSCM — пауч 150 г","Hot and Spicy Rice",150,1.07028522316359,0.938242888902583,"BSCM","Пауч"),
 ("BSCM — стакан 150 г","Thai Jasmine Rice",150,0.881437262584171,0.776655999619468,"BSCM","Чаша"),
 ("BSCM — стакан 150 г","Thai Brown Jasmine Rice",150,0.881407859256575,0.776655999619468,"BSCM","Чаша"),
 ("BSCM — стакан 150 г","Thai Red Jasmine Rice",150,0.954858269166893,0.841377226443071,"BSCM","Чаша"),
 ("BSCM — стакан 150 г","Rice Berry",150,0.954858269166893,0.841377226443071,"BSCM","Чаша"),
 ("BSCM — стакан 150 г","Brown Jasmine + Red Quinoa",150,0.954858269166893,0.841377226443071,"BSCM","Чаша"),
 ("BSCM — подвійний стакан 2×125 г","Jasmine Rice без олії",250,1.6275694066699,1.43357202931587,"BSCM","Чаша"),
 ("BSCM — подвійний стакан 2×125 г","Jasmine Rice з олією",250,1.6275694066699,1.43357202931587,"BSCM","Чаша"),
 ("BSCM — подвійний стакан 2×125 г","Jasmine (Pathum) з олією",250,1.47960898899142,1.30324749350685,"BSCM","Чаша"),
 ("BSCM — подвійний стакан 2×125 г","Brown Jasmine з олією",250,1.6275694066699,1.43357202931587,"BSCM","Чаша"),
 ("BSCM — подвійний стакан 2×125 г","Brown Rice з олією",250,1.62761752873563,1.43357202931587,"BSCM","Чаша"),
 ("BSCM — подвійний стакан 2×125 г","Basmati Rice з олією",250,1.77552982434839,1.56389656512489,"BSCM","Чаша"),
 ("BSCM — подвійний стакан 2×125 г","Long Grain Rice з олією",250,1.39083273838433,1.22505277202144,"BSCM","Чаша"),
 ("CM Premium — пауч 250 г органік","Organic Garlic Fried Rice",250,1.13300100081699,1.01722856884058,"CM Premium","Пауч"),
 ("CM Premium — пауч 250 г органік","Organic Cilantro Lime Rice",250,1.13300100081699,1.01722856884058,"CM Premium","Пауч"),
 ("CM Premium — пауч 250 г органік","Organic Turmeric Basmati",250,1.13300100081699,1.01722856884058,"CM Premium","Пауч"),
 ("CM Premium — пауч 250 г органік","Organic Jasmine Rice",250,1.13300100081699,1.01722856884058,"CM Premium","Пауч"),
 ("ONE'S / KBROS — лоток","Kimchi Fried Rice",200,2.1048630417902,1.93915406938722,"ONE'S","Лоток"),
 ("ONE'S / KBROS — лоток","Vegetable Fried Rice",200,2.1048630417902,1.93915406938722,"ONE'S","Лоток"),
 ("ONE'S / KBROS — лоток","Japchae Fried Rice",200,2.1048630417902,1.93915406938722,"ONE'S","Лоток"),
 ("ONE'S / KBROS — лоток","Cooked White Rice",210,0.971476418890301,0.894994727290563,"ONE'S","Лоток"),
]

# ---- known FOB (suppliers deck) ----
FOB = {
 ("BSCM — пауч 240 г","Jasmine Rice"):0.583, ("BSCM — пауч 240 г","Mexican Rice"):0.667,
 ("BSCM — пауч 240 г","Japanese Style Rice"):0.667, ("BSCM — пауч 240 г","Hot and Spicy Rice"):0.750,
 ("BSCM — пауч 200 г","Jasmine Rice (Pathum)"):0.550, ("BSCM — пауч 200 г","Mexican Rice"):0.620,
 ("BSCM — пауч 200 г","Japanese Style Rice"):0.620, ("BSCM — пауч 200 г","Hot and Spicy Rice"):0.680,
 ("BSCM — пауч 150 г","Jasmine Rice (Pathum)"):0.472, ("BSCM — пауч 150 г","Mexican Rice"):0.550,
 ("BSCM — пауч 150 г","Japanese Style Rice"):0.550, ("BSCM — пауч 150 г","Hot and Spicy Rice"):0.600,
 ("BSCM — стакан 150 г","Thai Jasmine Rice"):0.500, ("BSCM — стакан 150 г","Thai Brown Jasmine Rice"):0.500,
 ("CM Premium — пауч 250 г органік","Organic Jasmine Rice"):0.667,
 ("CM Premium — пауч 250 г органік","Organic Garlic Fried Rice"):0.667,
 ("CM Premium — пауч 250 г органік","Organic Cilantro Lime Rice"):0.667,
 ("CM Premium — пауч 250 г органік","Organic Turmeric Basmati"):0.667,
}

def shelf(cc_usd): return cc_usd * FX * SHELF_MULT
def partner(cc_usd): return cc_usd * FX / CC_SHARE

# ---- channel map (expanded) ----
CHANNELS = {
 "Мережевий роздріб (офлайн+онлайн)": {
   "found":["Сільпо"],
   "empty":["АТБ","Novus","METRO","Varus","Ашан","Fora","Таврія В","МегаМаркет","ЕКО маркет",
            "Ultramarket","Восторг","Космос","WineTime","Фуршет","Близенько","Точка","Рукавичка"],
   "note":"11 позицій, усі Ben's Original"},
 "Онлайн-супермаркети й маркетплейси": {
   "found":["MAUDAU","Rozetka","Prom","zakaz.ua (17 мереж)","GoToShop (агрегатор 8 мереж)"],
   "empty":[], "note":"Ben's · Henan · Bibigo · Clearspring · Portion"},
 "Азійські / етнічні фудшопи": {
   "found":["Смак Кореї","Тайякі Март","Pulsar","Апетітаріум","Gurmissimo","Edison Lee","OMG! Asia",
            "СНЕКІС","DCM","Скарби Азії","Sweet Svitt","Товари з Іспанії",
            "AsiaStyle","ASIA FOODS","Wèi Māo","AsianFoods","Asia Goods Store (Львів)",
            "SushiPovar","KladezGroup","Asia Foods Trade (опт)"],
   "empty":[], "note":"Ottogi · Henan · Bibigo · саморозігрів · Gallina Blanca"},
 "Туристичні та військові": {
   "found":["ALANTUR","Highlander","ForCamp","Гайдамака","Freeride","MK-Sport","Daruy","Tactico",
            "Лєєр","Activity","Desna","Шериф","OXO","Palmer","ВсеОпт","Суренж","Klever-Shop",
            "Terra Incognita","Висот-Нік","СУХПАЙ","UPcompany","SportStorm","Kalush-Craft",
            "110вольт","Клуб Мандрівник","Kamanti","Харчі ТМ","Їжа в Похід"],
   "empty":[], "note":"+ ще ~12 дрібних · Haidilao · Adventure Menu · усі сублімати"},
}
