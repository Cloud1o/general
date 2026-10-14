"""Dati demo condivisi: Giorgio Scala, due società immobiliari gemelle.

Valori di fantasia ma realistici (Campania). Le aliquote IMU sono esempi
da verificare sulla delibera del Comune (portale MEF - finanze.gov.it).
"""

ANNO = 2027

SOCIETA = [
    # codice, ragione sociale, saldo cassa 01/01, IRES+IRAP annua stimata,
    # assicurazioni annue, commercialista/consulenze mensili
    ("IU", "Immobiliare Uno S.r.l.", 85000, 18000, 6200, 650),
    ("ID", "Immobiliare Due S.r.l.", 60000, 12000, 4800, 550),
]

COMUNI = [
    # comune, codice catastale, aliquota anno prec. (per mille), aliquota anno corrente
    ("Napoli", "F839", 10.6, 10.6),
    ("Pozzuoli", "G964", 10.6, 10.6),
    ("Casoria", "B990", 10.6, 10.6),
    ("Caserta", "B963", 10.6, 10.6),
    ("Salerno", "H703", 10.4, 10.6),  # esempio di aumento deliberato dal Comune
]

# categoria catastale -> (moltiplicatore, gruppo, codice tributo quota Comune)
# Moltiplicatori L. 160/2019 c. 745; rendita rivalutata del 5%.
CATEGORIE = [
    ("A/1", 160, "A", "3918"), ("A/2", 160, "A", "3918"), ("A/3", 160, "A", "3918"),
    ("A/4", 160, "A", "3918"), ("A/7", 160, "A", "3918"), ("A/8", 160, "A", "3918"),
    ("A/10", 80, "A10", "3918"),
    ("B/1", 140, "B", "3918"),
    ("C/1", 55, "C1", "3918"), ("C/2", 160, "C", "3918"), ("C/3", 140, "C", "3918"),
    ("C/6", 160, "C", "3918"), ("C/7", 160, "C", "3918"),
    ("D/1", 65, "D", "3930"), ("D/5", 80, "D", "3930"), ("D/7", 65, "D", "3930"),
    ("D/8", 65, "D", "3930"),
]

# Occupazione 2027 per mese: 1 = locato, 0 = sfitto
TUTTO = [1] * 12

IMMOBILI = [
    # codice, società, descrizione, comune, tipo, categoria, rendita, valore mercato,
    # canone mensile, inquilino, mese anniversario contratto, % registro,
    # condominio mensile, manutenz. ordinaria annua (quota proprietà), bene merce,
    # occupazione 12 mesi
    ("IU-01", "IU", "Ufficio Centro Direzionale Is. E3", "Napoli", "Commerciale", "A/10",
     2450, 420000, 2300, "Studio Legale Ferraro", 3, 0.01, 180, 600, "Sì", TUTTO),
    ("IU-02", "IU", "Negozio Via Toledo", "Napoli", "Commerciale", "C/1",
     3900, 580000, 3800, "Bar Toledo S.n.c.", 9, 0.01, 120, 900, "Sì", TUTTO),
    ("IU-03", "IU", "Appartamento Vomero", "Napoli", "Abitativo", "A/2",
     1280, 390000, 1450, "Marco De Luca", 5, 0.02, 160, 500, "Sì", TUTTO),
    ("IU-04", "IU", "Appartamento Via Napoli", "Pozzuoli", "Abitativo", "A/3",
     780, 230000, 900, "Anna Esposito", 11, 0.02, 90, 400, "Sì", TUTTO),
    ("IU-05", "IU", "Box auto Via Napoli", "Pozzuoli", "Pertinenza", "C/6",
     160, 30000, 120, "Anna Esposito", 11, 0.02, 0, 50, "Sì", TUTTO),
    ("IU-06", "IU", "Capannone logistico", "Casoria", "Commerciale", "D/7",
     9800, 950000, 5500, "Logistica Campana S.p.A.", 1, 0.01, 0, 2500, "No", TUTTO),
    ("ID-01", "ID", "Appartamento Corso Trieste", "Caserta", "Abitativo", "A/2",
     1050, 260000, 1000, "Luca Russo", 6, 0.02, 110, 450, "Sì", TUTTO),
    ("ID-02", "ID", "Ufficio Via Roma", "Caserta", "Commerciale", "A/10",
     1900, 300000, 1700, "Assicurazioni Vesuvio S.r.l.", 4, 0.01, 140, 500, "Sì", TUTTO),
    ("ID-03", "ID", "Negozio Corso Vittorio Emanuele", "Salerno", "Commerciale", "C/1",
     2700, 410000, 2600, "Ottica Salernitana S.r.l.", 2, 0.01, 100, 700, "Sì", TUTTO),
    ("ID-04", "ID", "Appartamento Lungomare", "Salerno", "Abitativo", "A/3",
     900, 250000, 1100, "Giulia Ferri (disdetta 31/03)", 7, 0.02, 120, 450, "Sì",
     [1, 1, 1, 0, 0, 0, 0, 0, 1, 1, 1, 1]),
    ("ID-05", "ID", "Ufficio Centro Direzionale Is. G1", "Napoli", "Commerciale", "A/10",
     3100, 520000, 2900, "Tech Partners S.r.l.", 10, 0.01, 210, 650, "Sì", TUTTO),
    ("ID-06", "ID", "Posto auto Via Pergolesi", "Pozzuoli", "Pertinenza", "C/6",
     95, 18000, 80, "(sfitto)", 1, 0.02, 0, 0, "Sì", [0] * 12),
]

MUTUI = [
    # società, descrizione, banca, immobili a garanzia, debito residuo 01/01,
    # TAN attuale, rate residue, TAN ipotesi surroga
    ("IU", "Mutuo ipotecario Napoli", "Intesa Sanpaolo", "IU-01, IU-02", 250000, 0.038, 144, 0.031),
    ("IU", "Mutuo capannone", "BPER Banca", "IU-06", 380000, 0.041, 180, 0.034),
    ("ID", "Mutuo ipotecario Caserta/Salerno", "UniCredit", "ID-01, ID-02, ID-03", 190000, 0.035, 96, 0.030),
    ("ID", "Finanziamento chirografario", "BCC Campania", "-", 60000, 0.046, 30, 0.046),
]

# movimenti straordinari previsti: data, società, immobile, categoria, descrizione, importo (+ entrata / - uscita)
STRAORDINARI = [
    ("2027-04-15", "ID", "ID-04", "Manutenzione straordinaria", "Ritinteggiatura fine contratto", -2800),
    ("2027-05-10", "ID", "ID-04", "Manutenzione straordinaria", "Revisione completa impianti", -1900),
    ("2027-06-20", "ID", "ID-04", "Commerciale", "Annunci Immobiliare.it / Idealista", -350),
    ("2027-07-05", "IU", "IU-02", "Manutenzione straordinaria", "Sostituzione climatizzatore", -4200),
    ("2027-03-18", "IU", "IU-06", "Manutenzione straordinaria", "Rifacimento guaina copertura", -12500),
    ("2027-10-12", "ID", "ID-02", "Manutenzione straordinaria", "Infiltrazione infisso", -1100),
    ("2027-02-28", "ID", "ID-03", "Rateizzazione arretrati", "Ottica Salernitana - rata 1/6", 2600),
    ("2027-03-31", "ID", "ID-03", "Rateizzazione arretrati", "Ottica Salernitana - rata 2/6", 2600),
    ("2027-04-30", "ID", "ID-03", "Rateizzazione arretrati", "Ottica Salernitana - rata 3/6", 2600),
    ("2027-05-31", "ID", "ID-03", "Rateizzazione arretrati", "Ottica Salernitana - rata 4/6", 2600),
    ("2027-06-30", "ID", "ID-03", "Rateizzazione arretrati", "Ottica Salernitana - rata 5/6", 2600),
    ("2027-07-31", "ID", "ID-03", "Rateizzazione arretrati", "Ottica Salernitana - rata 6/6", 2600),
    ("2027-09-15", "IU", "-", "Compravendita", "Caparra acquisto nuovo immobile Caserta", -40000),
    ("2027-11-30", "ID", "ID-06", "Compravendita", "Vendita posto auto Via Pergolesi", 21000),
]
