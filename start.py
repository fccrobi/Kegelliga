import random
from typing import List, Any

# import xlwings
import numpy as np
import xlrd
from tabulate import tabulate


# a

# scheinbar erledigt??
# TODO Toni Zimmermann  119  129  130  121   499  2  0  1  2   497  121  139  123  114  Hagen Unger warum verliert Toni MP???

# TODO Kader werden aktuell nach Saisonende zurückgesetzt
# TODO ersten Spieler im Array transferieren können
# TODO Spielerwechsel geht nicht aber keine Fehlermeldung


# TODO Alter/Stärkeänderung
# TODO Ergebnisse als GUI mit farblichen Ergebnissen (Excel für anfang??)Spieltagsbericht als Excel exportieren, neue Tabelle für jeden Spieltag
# TODO Spieler und Stärken in Excel exportieren ("Speichern und LAden")
# TODO GUI Doppelklick auf Spieler mit Graph von Ergebnissen
# TODO Statistik Ergebnisse
# TODO Promotion / relegation (als erstes noch über manuelle Eingabe / später automatisch)
# TODO langsamer Anzeigemodus / Bahn für Bahn / Starter für Starter
# TODO Team managen und Austellung bestimmen
# TODO Bahnrekorde
# TODO Heimbahn Wählitz zb besser als Teuchern
# TODO Pokal
# TODO Geld für Transfers


class Liga:

    def __init__(self, anzahl, ligaebene, liganame):
        self.anzahl = anzahl
        self.ligaeben = ligaebene
        self.Ligaa = []
        self.Liganame = liganame
        self.Spielplan = []
        spieltag_nr = 1

        for i in range(0, sheet.nrows, 8):
            spieler = []
            if liganame == data[i][0]:
                for j in range(0, 8, 1):
                    spieler.append([data[i + j][3], data[i + j][4]])
                team_a = Verein(data[i][1], data[i][2], 0, spieler, 0, 0, 0, 0, 0, 0)
                self.Ligaa.append(team_a)
        spielplan = self.spielplanGenerator(anzahl)
        self.menu(spielplan, spieltag_nr)

    def menu(self, spielplan, spieltag_nr):
        while 1:
            print("")
            print("1 = Team hinzufügen")
            print("3 = Statistik Team")
            print("4 = Tabelle")
            print("5 = Spieltag")
            print("6 = Spielerwechsel")
            print("0 = Saison beenden")
            print("")
            inp: str = input()
            if inp == "1":
                self.Teamadd()
            elif inp == "3":
                self.Statistik()
            elif inp == "4":
                self.Tabelle(spieltag_nr)
            elif inp == "5":
                vorbei = self.Spieltag(spielplan, spieltag_nr)
                if vorbei == 2:
                    spieltag_nr -= 1
                spieltag_nr += 1
            elif inp == "6":
                self.Spielerwechsel()
            elif inp == "0":
                print("Bitte nochmal mit 0 bestätigen!")
                inp = input()
                if inp == "0":
                    break

    # benötigt für Spielplan
    def make_day(self, num_teams, day):
        # generate list of teams
        lst = list(range(1, num_teams + 1))
        # rotate
        day %= (num_teams - 1)  # clip to 0 .. num_teams - 2
        if day:  # if day == 0, no rotation is needed (and using -0 as list index will cause problems)
            lst = lst[:1] + lst[-day:] + lst[1:-day]
        half = num_teams // 2
        return list(zip(lst[:half], lst[half:][::-1]))

    # benötigt für Spielplan
    def spielplanGenerator(self, Anzahl):
        # Gerade Anzahl
        if Anzahl % 2:
            Anzahl += 1

        schedule = [self.make_day(Anzahl, day) for day in range(Anzahl - 1)]
        # Tauch Heim und Auswärts
        swapped = [[(away, home) for home, away in day] for day in schedule]
        return schedule + swapped

    # ohne Funktion
    def Teamadd(self):
        if (Liga == self.anzahl):
            print("Liga voll!")

    # TeamStatistik ausgeben
    def Statistik(self):
        print("Welcher Verein?")
        inp = input()
        for Verein in self.Ligaa:
            if (Verein.Name == inp):
                print(Verein.Name)
                print(Verein.Stärke)
                print(Verein.Punkte)
                return
        print("Verein nicht gefunden")

    # Tabelle ausgeben
    def Tabelle(self, SpieltagNr):
        self.Ligaa.sort(key=lambda Verein: Verein.Punkte, reverse=1)
        Kopie = []
        for Verein in self.Ligaa:
            try:
                Kopie.append([Verein.Name, Verein.Punkte, Verein.S, Verein.U, Verein.N, Verein.MP, Verein.SP,
                              Verein.Schnitt / (SpieltagNr - 1)])
            except:
                pass
        print(tabulate(Kopie, headers=["Name", "Punkte", "S", "U", "N", "MP", "SP", "Schnitt"]))

    def Spieltag(self, Spielplan, SpieltagNr):
        print("Spieltag Nr: " + str(SpieltagNr))
        # prüfen ob Saison vorbei ist
        if (len(Spielplan) + 1 <= SpieltagNr):
            print("Saison bereits vorbei")
            return 2
        else:
            # for Schleife für Anzahl Spiele am Spieltag
            for i in range(0, len(Spielplan[SpieltagNr - 1]), 1):
                try:
                    Spiel2 = Spiel(self.Ligaa[Spielplan[SpieltagNr - 1][i][0] - 1],
                                   self.Ligaa[Spielplan[SpieltagNr - 1][i][1] - 1])
                except:
                    print("Anzahl Teams Spielplan und Anzahl Teams Liga stimmen nicht überein")
                    return

    def Spielerwechsel(self):

        print("Spieler 1:")
        Sp1 = input()
        i1 = 0
        j1 = 0
        # Position im Array Spieler 1 finden
        for i in range(0, len(self.Ligaa), 1):
            for j in range(0, len(self.Ligaa[i].Spieler), 1):
                if (self.Ligaa[i].Spieler[j][0] == Sp1):
                    i1 = i
                    j1 = j
                    break
        if (i1 == 0 and j1 == 0):
            print("Spieler nicht gefunden")
            return 0

        print("Spieler 2:")
        Sp2 = input()
        i2 = 0
        j2 = 0
        # Position im Array Spieler 1 finden
        for i in range(0, len(self.Ligaa), 1):
            for j in range(0, len(self.Ligaa[i].Spieler), 1):
                if (self.Ligaa[i].Spieler[j][0] == Sp2):
                    i2 = i
                    j2 = j
                    break
        if (i2 == 0 and j2 == 0):
            print("Spieler nicht gefunden")
            return 0

        self.Ligaa[i1].Spieler.append(self.Ligaa[i2].Spieler[j2])
        self.Ligaa[i2].Spieler.append(self.Ligaa[i1].Spieler[j1])
        self.Ligaa[i1].Spieler.pop(j1)
        self.Ligaa[i2].Spieler.pop(j2)


class Verein:
    def __init__(self, Name, stärke, punkte, Spieler, MP, SP, Schnitt, S, U, N):
        self.Name = Name
        self.Stärke = stärke
        self.Punkte = 0
        self.Spieler = Spieler
        self.MP = MP
        self.SP = SP
        self.Schnitt = Schnitt
        self.S = S
        self.U = U
        self.N = N

    def sieg(self, MP, SP, Schnitt):
        self.Punkte = self.Punkte + 3
        self.MP = self.MP + MP
        self.SP = self.SP + SP
        self.Schnitt = self.Schnitt + Schnitt
        self.S = self.S + 1

    def unentschieden(self, mp, sp, Schnitt):
        self.Punkte = self.Punkte + 1
        self.MP = self.MP + mp
        self.SP = self.SP + sp
        self.Schnitt = self.Schnitt + Schnitt
        self.U = self.U + 1

    def niederlage(self, MP, SP, Schnitt):
        self.MP = self.MP + MP
        self.SP = self.SP + SP
        self.Schnitt = self.Schnitt + Schnitt
        self.N = self.N + 1


class Spiel:
    def __init__(self, TeamA, TeamB):
        self.TeamA = TeamA
        self.TeamB = TeamB

        print(TeamA.Name + " - " + TeamB.Name)

        # 2 zufälige Spieler nicht da
        # TeamA
        Ersatz: List[Any] = []
        rd = random.randrange(0, 8)
        Ersatz.append(TeamA.Spieler[rd])
        TeamA.Spieler.pop(rd)
        rd = random.randrange(0, 7)
        Ersatz.append(TeamA.Spieler[rd])
        TeamA.Spieler.pop(rd)
        # TeamB
        ersatz2 = []
        rd = random.randrange(0, 8)
        ersatz2.append(TeamB.Spieler[rd])
        TeamB.Spieler.pop(rd)
        rd = random.randrange(0, 7)
        ersatz2.append(TeamB.Spieler[rd])
        TeamB.Spieler.pop(rd)

        Ergeb = [[(0) for c in range(0, 16)] for r in range(0, 7)]
        for j in range(0, 6, 1):
            # Namen eintragen
            Ergeb[j][0] = TeamA.Spieler[j][0]
            Ergeb[j][15] = TeamB.Spieler[j][0]
            # Tagesform, beeinflusst Gesamtergebnis Spieler
            HrandA = np.random.normal(1000, 30, 1)
            GrandA = np.random.normal(1000, 30, 1)
            for i in range(1, 5, 1):
                Hrand = np.random.normal(1000, 70, 1)
                Ergeb[j][i] = int(TeamA.Spieler[j][1] * (HrandA / 1000) * (Hrand / 1000) / 4)
                # Ergebnis Spieler
                Ergeb[j][5] += Ergeb[j][i]

                Grand = np.random.normal(1000, 70, 1)
                Ergeb[j][i + 10] = int(TeamB.Spieler[j][1] * (GrandA / 1000) * (Grand / 1000) / 4)
                # Ergebnis Spieler
                Ergeb[j][10] += Ergeb[j][i + 10]
                # SP Spieler
                if Ergeb[j][i + 10] > Ergeb[j][i]:
                    Ergeb[j][9] += 1
                elif Ergeb[j][i + 10] == Ergeb[j][i]:
                    Ergeb[j][9] += 0.5
                # MP Spieler
                if Ergeb[j][9] > 2:
                    Ergeb[j][8] = 1
                elif Ergeb[j][9] == 2:
                    if Ergeb[j][10] > Ergeb[j][5]:
                        Ergeb[j][8] = 1
                    elif Ergeb[j][10] == Ergeb[j][5]:
                        Ergeb[j][8] = 0.5

            # Punkte Heim
            # SP Spieler
            Ergeb[j][6] = 4 - Ergeb[j][9]
            # MP Spieler
            Ergeb[j][7] = 1 - Ergeb[j][8]
            # Gesamtholz Mannschaft
            Ergeb[6][5] += Ergeb[j][5]
            Ergeb[6][10] += Ergeb[j][10]
            # MP Gesamt
            Ergeb[6][7] += Ergeb[j][7]
            Ergeb[6][8] += Ergeb[j][8]
            # SP Gesa,t
            Ergeb[6][6] += Ergeb[j][6]
            Ergeb[6][9] += Ergeb[j][9]
        # MP für Gesamtholz
        if Ergeb[6][5] > Ergeb[6][10]:
            Ergeb[6][7] += 2
        elif Ergeb[6][5] == Ergeb[6][10]:
            Ergeb[6][7] += 1
            Ergeb[6][8] += 1
        else:
            Ergeb[6][8] += 2

        print(tabulate(Ergeb,
                       headers=["Name", "B1", "B2", "B3", "B4", "G", "SP", "MP", "MP", "SP", "G", "B1", "B2", "B3",
                                "B4", "Name"]))
        # Tabellenpunkte
        if Ergeb[6][7] > Ergeb[6][8]:
            TeamA.sieg(Ergeb[6][7], Ergeb[6][6], Ergeb[6][5])
            TeamB.niederlage(Ergeb[6][8], Ergeb[6][9], Ergeb[6][10])
        elif Ergeb[6][7] == Ergeb[6][8]:
            TeamA.unentschieden(Ergeb[6][7], Ergeb[6][6], Ergeb[6][5])
            TeamB.unentschieden(Ergeb[6][8], Ergeb[6][9], Ergeb[6][10])
        else:
            TeamB.sieg(Ergeb[6][8], Ergeb[6][9], Ergeb[6][10])
            TeamA.niederlage(Ergeb[6][7], Ergeb[6][6], Ergeb[6][5])

        # nicht bereite Spieler wieder anhängen
        TeamA.Spieler.append(Ersatz[0])
        TeamA.Spieler.append(Ersatz[1])
        TeamB.Spieler.append(ersatz2[0])
        TeamB.Spieler.append(ersatz2[1])


book = xlrd.open_workbook('Input.xlsx')
sheet = book.sheet_by_name('Tabelle1')
data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

while 1:
    print("")
    print("1 = Saison starten")
    print("")
    inp: str = input()
    if inp == "1":
        Anzahl = 14
        Kreisoberliga = Liga(Anzahl, 1, "Kreisoberliga")
