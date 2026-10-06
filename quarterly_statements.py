"""Quarterly statement lines for the DuPont, the comparison, and the forecast.

USD millions. Pulled from SEC companyfacts XBRL for the Form 10-Q and Form 10-K
quarters from Q2 2024 through Q2 2026. A tagged quarter is used when the filing
has one. The fourth quarter, and a cash-flow quarter that is tagged only as a
year-to-date total, is the year-to-date total minus the prior year-to-date total.
Q2 24 is the opening balance. It is not a ratio quarter.

Intel net income is consolidated profit, the line that ties to the annual total
of $26m in FY25. Intel equity includes non-controlling interests. AMD interest
income is annual-only in the filing taxonomy, so ii is None and modified DuPont
is not computed for AMD.
"""

from __future__ import annotations

TAX = 0.21
SLOTS = ['Q3 24', 'Q4 24', 'Q1 25', 'Q2 25', 'Q3 25', 'Q4 25', 'Q1 26', 'Q2 26']
ORDER = ('NVIDIA', 'AMD', 'Intel')

SHARES = {'NVIDIA': {'diluted_m': 24285.0, 'outstanding_m': None}, 'AMD': {'diluted_m': 1659.0, 'outstanding_m': 1632.0}, 'Intel': {'diluted_m': 5104.0, 'outstanding_m': 5043.0}}

QUARTERS = {
    'NVIDIA': [
        {'end': '2024-07-28', 'slot': 'Q2 24', 'rev': 30040, 'oi': 18642, 'ni': 16599, 'ix': 61, 'ii': 444, 'ocf': 14488, 'capex': 977, 'ta': 85227, 'eq': 58157, 'cash': 8563, 'liq': 26237, 'lease': 1304, 'debt': 8461},
        {'end': '2024-10-27', 'slot': 'Q3 24', 'rev': 35082, 'oi': 21869, 'ni': 19309, 'ix': 61, 'ii': 472, 'ocf': 17627, 'capex': 813, 'ta': 96013, 'eq': 65899, 'cash': 9107, 'liq': 29380, 'lease': 1490, 'debt': 8462},
        {'end': '2025-01-26', 'slot': 'Q4 24', 'rev': 39331, 'oi': 24033, 'ni': 22091, 'ix': 61, 'ii': 511, 'ocf': 16629, 'capex': 1077, 'ta': 111601, 'eq': 79327, 'cash': 8589, 'liq': 34621, 'lease': 1519, 'debt': 8463},
        {'end': '2025-04-27', 'slot': 'Q1 25', 'rev': 44062, 'oi': 21638, 'ni': 18775, 'ix': 63, 'ii': 515, 'ocf': 27414, 'capex': 1227, 'ta': 125254, 'eq': 83843, 'cash': 15234, 'liq': 38457, 'lease': 1521, 'debt': 8464},
        {'end': '2025-07-27', 'slot': 'Q2 25', 'rev': 46743, 'oi': 28440, 'ni': 26422, 'ix': 62, 'ii': 592, 'ocf': 15365, 'capex': 1895, 'ta': 140740, 'eq': 100131, 'cash': 11639, 'liq': 45152, 'lease': 1831, 'debt': 8466},
        {'end': '2025-10-26', 'slot': 'Q3 25', 'rev': 57006, 'oi': 36010, 'ni': 31910, 'ix': 61, 'ii': 624, 'ocf': 23751, 'capex': 1636, 'ta': 161148, 'eq': 118897, 'cash': 11486, 'liq': 49122, 'lease': 2014, 'debt': 8467},
        {'end': '2026-01-25', 'slot': 'Q4 25', 'rev': 68127, 'oi': 44299, 'ni': 42960, 'ix': 73, 'ii': 569, 'ocf': 36188, 'capex': 1284, 'ta': 206803, 'eq': 157293, 'cash': 10605, 'liq': 51951, 'lease': 2572, 'debt': 8468},
        {'end': '2026-04-26', 'slot': 'Q1 26', 'rev': 81615, 'oi': 53536, 'ni': 58321, 'ix': 102, 'ii': 540, 'ocf': 50344, 'capex': 1757, 'ta': 259474, 'eq': 195474, 'cash': 13237, 'liq': 67335, 'lease': 3878, 'debt': 8470},
        {'end': '2026-07-26', 'slot': 'Q2 26', 'rev': 96221, 'oi': 63734, 'ni': 59688, 'ix': 227, 'ii': 496, 'ocf': 24077, 'capex': 2677, 'ta': 320272, 'eq': 228984, 'cash': 22443, 'liq': 76926, 'lease': 4985, 'debt': 33366},
    ],
    'AMD': [
        {'end': '2024-06-29', 'slot': 'Q2 24', 'rev': 5835, 'oi': 269, 'ni': 265, 'ix': 25, 'ii': None, 'ocf': 593, 'capex': 154, 'ta': 67886, 'eq': 56538, 'cash': 4113, 'liq': 1227, 'lease': 526, 'debt': 1719},
        {'end': '2024-09-28', 'slot': 'Q3 24', 'rev': 6819, 'oi': 724, 'ni': 771, 'ix': 23, 'ii': None, 'ocf': 628, 'capex': 132, 'ta': 69636, 'eq': 56985, 'cash': 3897, 'liq': 647, 'lease': 518, 'debt': 1720},
        {'end': '2024-12-28', 'slot': 'Q4 24', 'rev': 7658, 'oi': 871, 'ni': 482, 'ix': 19, 'ii': None, 'ocf': 1299, 'capex': 208, 'ta': 69226, 'eq': 57568, 'cash': 3787, 'liq': 1345, 'lease': 491, 'debt': 1721},
        {'end': '2025-03-29', 'slot': 'Q1 25', 'rev': 7438, 'oi': 806, 'ni': 709, 'ix': 20, 'ii': None, 'ocf': 939, 'capex': 212, 'ta': 71550, 'eq': 57881, 'cash': 6049, 'liq': 1261, 'lease': 567, 'debt': 4164},
        {'end': '2025-06-28', 'slot': 'Q2 25', 'rev': 7685, 'oi': -134, 'ni': 872, 'ix': 38, 'ii': None, 'ocf': 2011, 'capex': 282, 'ta': 74820, 'eq': 59665, 'cash': 4442, 'liq': 1425, 'lease': 668, 'debt': 3218},
        {'end': '2025-09-27', 'slot': 'Q3 25', 'rev': 9246, 'oi': 1270, 'ni': 1243, 'ix': 37, 'ii': None, 'ocf': 2159, 'capex': 258, 'ta': 76891, 'eq': 60790, 'cash': 4808, 'liq': 2435, 'lease': 650, 'debt': 3220},
        {'end': '2025-12-27', 'slot': 'Q4 25', 'rev': 10270, 'oi': 1752, 'ni': 1511, 'ix': 36, 'ii': None, 'ocf': 2600, 'capex': 222, 'ta': 76926, 'eq': 62999, 'cash': 5539, 'liq': 5013, 'lease': 625, 'debt': 3222},
        {'end': '2026-03-28', 'slot': 'Q1 26', 'rev': 10253, 'oi': 1476, 'ni': 1383, 'ix': 37, 'ii': None, 'ocf': 2955, 'capex': 389, 'ta': 79642, 'eq': 64462, 'cash': 5585, 'liq': 6762, 'lease': 647, 'debt': 3224},
        {'end': '2026-06-27', 'slot': 'Q2 26', 'rev': 11536, 'oi': 1990, 'ni': 2297, 'ix': 37, 'ii': None, 'ocf': 2366, 'capex': 808, 'ta': 84464, 'eq': 67224, 'cash': 5086, 'liq': 8025, 'lease': 1050, 'debt': 3226},
    ],
    'Intel': [
        {'end': '2024-06-29', 'slot': 'Q2 24', 'rev': 12833, 'oi': -1964, 'ni': -1654, 'ix': 294, 'ii': 320, 'ocf': 2292, 'capex': 5682, 'ta': 206205, 'eq': 120434, 'cash': 11287, 'liq': 17986, 'lease': 0, 'debt': 53029},
        {'end': '2024-09-28', 'slot': 'Q3 24', 'rev': 13284, 'oi': -9057, 'ni': -16989, 'ix': 248, 'ii': 340, 'ocf': 4054, 'capex': 6458, 'ta': 193542, 'eq': 104864, 'cash': 8785, 'liq': 15301, 'lease': 0, 'debt': 50236},
        {'end': '2024-12-28', 'slot': 'Q4 24', 'rev': 14260, 'oi': 412, 'ni': -153, 'ix': 234, 'ii': 262, 'ocf': 3165, 'capex': 5834, 'ta': 196485, 'eq': 105032, 'cash': 8249, 'liq': 13813, 'lease': 0, 'debt': 50011},
        {'end': '2025-03-29', 'slot': 'Q1 25', 'rev': 12667, 'oi': -301, 'ni': -887, 'ix': 299, 'ii': 245, 'ocf': 813, 'capex': 5183, 'ta': 192242, 'eq': 106413, 'cash': 8947, 'liq': 12101, 'lease': 0, 'debt': 50151},
        {'end': '2025-06-28', 'slot': 'Q2 25', 'rev': 12859, 'oi': -3176, 'ni': -3024, 'ix': 227, 'ii': 210, 'ocf': 2050, 'capex': 3550, 'ta': 192520, 'eq': 105751, 'cash': 9643, 'liq': 11563, 'lease': 0, 'debt': 50757},
        {'end': '2025-09-27', 'slot': 'Q3 25', 'rev': 13653, 'oi': 683, 'ni': 4270, 'ix': 282, 'ii': 228, 'ocf': 2546, 'capex': 2425, 'ta': 204514, 'eq': 116730, 'cash': 11141, 'liq': 19794, 'lease': 0, 'debt': 46553},
        {'end': '2025-12-27', 'slot': 'Q4 25', 'rev': 13674, 'oi': 580, 'ni': -333, 'ix': 283, 'ii': 324, 'ocf': 4288, 'capex': 3488, 'ta': 211429, 'eq': 126360, 'cash': 14265, 'liq': 23151, 'lease': 0, 'debt': 46585},
        {'end': '2026-03-28', 'slot': 'Q1 26', 'rev': 13577, 'oi': -3136, 'ni': -4281, 'ix': 264, 'ii': 333, 'ocf': 1096, 'capex': 3636, 'ta': 205332, 'eq': 124989, 'cash': 17247, 'liq': 15542, 'lease': 0, 'debt': 45031},
        {'end': '2026-06-27', 'slot': 'Q2 26', 'rev': 16128, 'oi': 1796, 'ni': -10848, 'ix': 321, 'ii': 334, 'ocf': 7006, 'capex': 2556, 'ta': 202439, 'eq': 103143, 'cash': 12874, 'liq': 16853, 'lease': 0, 'debt': 50537},
    ],
}

