class Document:
    def __init__(self, boekjaar, dagboek, nummer, datum, relatiecode, vervaldatum, tebetalen, basisbedrag):
        self.boekjaar = boekjaar
        self.dagboek = dagboek
        self.nummer = nummer
        self.periode = datum.strftime("%m")
        self.btwregime = "H"
        self.datum = datum.strftime("%d/%m/%Y")
        self.relatiecode = relatiecode
        self.vervaldatum = vervaldatum.strftime("%d/%m/%Y")
        self.tebetalen = tebetalen
        self.basisbedrag = basisbedrag
        self.btwtebetalen = tebetalen - basisbedrag
        self.btwcode = self.get_btw_code()
        self.valutacode = "EUR"
        self.omschrijving = ""  # Placeholder
        self.koers = ""  # Placeholder
        self.vertegenw_code = ""  # Placeholder
        self.statusfact_status = "OK"  # Placeholder
        self.codefcbd_codefcbd = "F"
        self.kortcont = ""  # Placeholder
        self.bedragkc = ""  # Placeholder
        self.betvoorw_code = ""  # Placeholder
        self.ogm = ""  # Placeholder
        self.kredbep = ""  # Placeholder
        self.driehoeksverkeer = ""  # Placeholder
        self.boekhpl_reknr = 700002  # Placeholder
        self.omschr_d = ""  # Placeholder
        self.datum_d = self.datum  # Placeholder
        self.bedrag_d = self.basisbedrag  # Placeholder
        self.btwcodes_btwcode_d = self.btwcode  # Placeholder
        self.code_a = ""  # Placeholder
        self.bedrag_a = ""  # Placeholder

    def get_btw_code(self):
        if self.basisbedrag == 0:
            return 0
        else:
            ratio = round(self.tebetalen / self.basisbedrag, 2)
            if ratio == 1.00:
                return 5
            elif ratio == 1.06:
                return 2
            elif ratio == 1.12:
                return 3
            elif ratio == 1.21:
                return 4
            else:
                return "FOUT"

    def to_dict(self):
        return {
            'boekjaar_boekjaar (H)': self.boekjaar,
            'dagboek_dagboek (H)': self.dagboek,
            'factuur (H)': self.nummer,
            'periode_periode (H)': self.periode,
            'btwregimes_btwregime (H)': self.btwregime,
            'factdat (H)': self.datum,
            'relaties_code (H)': self.relatiecode,
            'vervdat (H)': self.vervaldatum,
            'omschrijving (H)': self.omschrijving,
            'valuta_code (H)': self.valutacode,
            'koers (H)': self.koers,
            'vertegenw_code (H)': self.vertegenw_code,
            'tebet (H)': self.tebetalen,
            'statusfact_status (H)': self.statusfact_status,
            'codefcbd_codefcbd (H)': self.codefcbd_codefcbd,
            'kortcont (H)': self.kortcont,
            'bedragkc (H)': self.bedragkc,
            'basis (H)': self.basisbedrag,
            'btwtebet (H)': self.btwtebetalen,
            'betvoorw_code (H)': self.betvoorw_code,
            'ogm (H)': self.ogm,
            'kredbep (H)': self.kredbep,
            'driehoeksverkeer (H)': self.driehoeksverkeer,
            'boekhpl_reknr (D)': self.boekhpl_reknr,
            'omschr (D)': self.omschr_d,
            'datum (D)': self.datum_d,
            'bedrag (D)': self.bedrag_d,
            'btwcodes_btwcode (D)': self.btwcodes_btwcode_d,
            'code (A)': self.code_a,
            'bedrag (A)': self.bedrag_a,
        }

