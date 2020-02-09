import random
import sys
from typing import List, Any

import numpy as np
import xlrd
import xlwings as xw
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QApplication, QWidget, QTableWidget, QTableWidgetItem, QVBoxLayout, QHeaderView
from tabulate import tabulate


# scheinbar erledigt??
# TODO Toni Zimmermann  119  129  130  121   499  2  0  1  2   497  121  139  123  114  Hagen Unger warum verliert Toni MP???

# TODO Ergebnisse als GUI mit farblichen Ergebnissen
# TODO GUI Doppelklick auf Spieler mit Graph von Ergebnissen
# TODO Alter/Stärkeänderung, Neugeneration Spieler, Normalverteilung nach Alter????
# TODO Formkurve Spieler, alle paar Woche mal schlechte Form z.B.
# TODO Saisoncounter, jede Saison neues Tabellenblatt mit Rostern und Tabellen
# TODO Statistik Ergebnisse
# TODO mehrere Klassen Promotion / relegation (als erstes noch über manuelle Eingabe / später automatisch)
# TODO langsamer Anzeigemodus / Bahn für Bahn / Starter für Starter
# TODO Team managen und Austellung bestimmen
# TODO Bahnrekorde
# TODO Heimbahn Wählitz zb besser als Teuchern
# TODO Pokal
# TODO Geld für Transfers
# TODO Neurales Netzwerk lernt wass es machen muss um erster in Tabelle zu werden, vorher umstellen Spielerwechsel über Name zu Spielerwechsel nach i???


class Liga:

    def __init__(self, anzahl, ligaebene, liganame):
        self.anzahl = anzahl
        self.ligaeben = ligaebene
        self.Ligaa = []
        self.Liganame = liganame
        self.Spielplan = []
        spieltag_nr = 1

        # aus data zu Objekten
        for i in range(0, sheet.nrows, 8):
            spieler = []
            if liganame == data[i][0]:
                for j in range(0, 8, 1):
                    spieler.append([data[i + j][3], data[i + j][4]])
                team_a = Verein(data[i][1], data[i][2], spieler, 0, 0, 0, 0, 0, 0)
                self.Ligaa.append(team_a)
        spielplan = self.spielplanGenerator(anzahl)
        self.menu(spielplan, spieltag_nr)

        # aus Objekten zu Excel
        wb = xw.Book("Input.xlsx")
        sht = xw.sheets("Tabelle1")
        print("Moment,speichern.....")
        for i in range(0, len(self.Ligaa) * 8, 8):
            sht.range((i + 1, 1)).value = self.Liganame
            sht.range((i + 1, 2)).value = self.Ligaa[int(i / 8)].Name
            for j in range(0, 8, 1):
                sht.range((i + j + 1, 4)).value = self.Ligaa[int(i / 8)].Spieler[j]
        wb.save()  # speichern, da sonst Änderungen nicht geladen werden

    def menu(self, spielplan, spieltag_nr):
        while 1:
            print("")
            print("1 = Team hinzufügen")
            print("3 = Statistik Team")
            print("4 = Tabelle")
            print("5 = Spieltag")
            print("6 = Spielerwechsel")
            print("0 = Saison beenden und Teams speichern")
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
            if Verein.Name == inp:
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
        # Tabelle-GUI
        app = QApplication(sys.argv)
        Tabelle = Table(Kopie)
        # Wichtige Zeile, damit Fenster offen bleibt, aber nach Schließung das Programm weiterläuft
        app.exec_()

    def Spieltag(self, Spielplan, SpieltagNr):
        Spieltagserg = []
        print("Spieltag Nr: " + str(SpieltagNr))
        # prüfen ob Saison vorbei ist
        if (len(Spielplan) + 1 <= SpieltagNr):
            print("Saison bereits vorbei")
            return 2
        else:
            wb = xw.Book("Output.xlsx")
            try:
                wb.sheets[str(SpieltagNr)]
            except:
                wb.sheets.add(str(SpieltagNr))
            zeile = 1

            # for Schleife für Anzahl Spiele am Spieltag
            for i in range(0, len(Spielplan[SpieltagNr - 1]), 1):
                try:
                    Spiel2 = Spiel(self.Ligaa[Spielplan[SpieltagNr - 1][i][0] - 1],
                                   self.Ligaa[Spielplan[SpieltagNr - 1][i][1] - 1], zeile, SpieltagNr)
                    zeile += 10
                    wb.save()
                except:
                    print("Anzahl Teams Spielplan und Anzahl Teams Liga stimmen nicht überein")
                    return

    # Tabelle-GUI
    # app = QApplication(sys.argv)
    # Spieltagsuebersicht = Spieltagsuebersicht(Kopie)
    # Wichtige Zeile, damit Fenster offen bleibt, aber nach Schließung das Programm weiterläuft
    #app.exec_()


    def Spielerwechsel(self):

        print("Spieler 1:")
        Sp1 = input()
        i1 = -1
        j1 = 0
        # Position im Array Spieler 1 finden
        for i in range(0, len(self.Ligaa), 1):
            for j in range(0, len(self.Ligaa[i].Spieler), 1):
                if (self.Ligaa[i].Spieler[j][0] == Sp1):
                    i1 = i
                    j1 = j
                    break
        if (i1 == -1 and j1 == 0):
            print("Spieler nicht gefunden")
            return 0

        print("Spieler 2:")
        Sp2 = input()
        i2 = -1
        j2 = 0
        # Position im Array Spieler 1 finden
        for i in range(0, len(self.Ligaa), 1):
            for j in range(0, len(self.Ligaa[i].Spieler), 1):
                if (self.Ligaa[i].Spieler[j][0] == Sp2):
                    i2 = i
                    j2 = j
                    break
        if (i2 == -1 and j2 == 0):
            print("Spieler nicht gefunden")
            return 0

        self.Ligaa[i1].Spieler.append(self.Ligaa[i2].Spieler[j2])
        self.Ligaa[i2].Spieler.append(self.Ligaa[i1].Spieler[j1])
        self.Ligaa[i1].Spieler.pop(j1)
        self.Ligaa[i2].Spieler.pop(j2)


class Verein:
    def __init__(self, name, stärke, spieler, mp, sp, schnitt, s, u, n):
        self.Name = name
        self.Stärke = stärke
        self.Punkte = 0
        self.Spieler = spieler
        self.MP = mp
        self.SP = sp
        self.Schnitt = schnitt
        self.S = s
        self.U = u
        self.N = n

    def sieg(self, mp, sp, schnitt):
        self.Punkte = self.Punkte + 3
        self.MP = self.MP + mp
        self.SP = self.SP + sp
        self.Schnitt = self.Schnitt + schnitt
        self.S = self.S + 1

    def unentschieden(self, mp, sp, schnitt):
        self.Punkte = self.Punkte + 1
        self.MP = self.MP + mp
        self.SP = self.SP + sp
        self.Schnitt = self.Schnitt + schnitt
        self.U = self.U + 1

    def niederlage(self, mp, sp, schnitt):
        self.MP = self.MP + mp
        self.SP = self.SP + sp
        self.Schnitt = self.Schnitt + schnitt
        self.N = self.N + 1


class Spiel:
    def __init__(self, team_a, team_b, zeile, SpieltagNr):
        self.TeamA = team_a
        self.TeamB = team_b
        self.zeile = zeile
        self.SpieltagNr = SpieltagNr



        print(team_a.Name + " - " + team_b.Name)

        # 2 zufälige Spieler nicht da
        # team_a
        ersatz: List[Any] = []
        rd = random.randrange(0, 8)
        ersatz.append(team_a.Spieler[rd])
        team_a.Spieler.pop(rd)
        rd = random.randrange(0, 7)
        ersatz.append(team_a.Spieler[rd])
        team_a.Spieler.pop(rd)
        # team_b
        ersatz2 = []
        rd = random.randrange(0, 8)
        ersatz2.append(team_b.Spieler[rd])
        team_b.Spieler.pop(rd)
        rd = random.randrange(0, 7)
        ersatz2.append(team_b.Spieler[rd])
        team_b.Spieler.pop(rd)

        ergeb = [[(0) for c in range(0, 16)] for r in range(0, 7)]
        for j in range(0, 6, 1):
            # Namen eintragen
            ergeb[j][0] = team_a.Spieler[j][0]
            ergeb[j][15] = team_b.Spieler[j][0]
            # Tagesform, beeinflusst Gesamtergebnis Spieler
            hrand_a = np.random.normal(1000, 30, 1)
            grand_a = np.random.normal(1000, 30, 1)
            for i in range(1, 5, 1):
                hrand = np.random.normal(1000, 70, 1)
                ergeb[j][i] = int(team_a.Spieler[j][1] * (hrand_a / 1000) * (hrand / 1000) / 4)
                # Ergebnis Spieler
                ergeb[j][5] += ergeb[j][i]

                grand = np.random.normal(1000, 70, 1)
                ergeb[j][i + 10] = int(team_b.Spieler[j][1] * (grand_a / 1000) * (grand / 1000) / 4)
                # Ergebnis Spieler
                ergeb[j][10] += ergeb[j][i + 10]
                # SP Spieler
                if ergeb[j][i + 10] > ergeb[j][i]:
                    ergeb[j][9] += 1
                elif ergeb[j][i + 10] == ergeb[j][i]:
                    ergeb[j][9] += 0.5
                # MP Spieler
                if ergeb[j][9] > 2:
                    ergeb[j][8] = 1
                elif ergeb[j][9] == 2:
                    if ergeb[j][10] > ergeb[j][5]:
                        ergeb[j][8] = 1
                    elif ergeb[j][10] == ergeb[j][5]:
                        ergeb[j][8] = 0.5

            # Punkte Heim
            # SP Spieler
            ergeb[j][6] = 4 - ergeb[j][9]
            # MP Spieler
            ergeb[j][7] = 1 - ergeb[j][8]
            # Gesamtholz Mannschaft
            ergeb[6][5] += ergeb[j][5]
            ergeb[6][10] += ergeb[j][10]
            # MP Gesamt
            ergeb[6][7] += ergeb[j][7]
            ergeb[6][8] += ergeb[j][8]
            # SP Gesa,t
            ergeb[6][6] += ergeb[j][6]
            ergeb[6][9] += ergeb[j][9]
        # MP für Gesamtholz
        if ergeb[6][5] > ergeb[6][10]:
            ergeb[6][7] += 2
        elif ergeb[6][5] == ergeb[6][10]:
            ergeb[6][7] += 1
            ergeb[6][8] += 1
        else:
            ergeb[6][8] += 2

        print(tabulate(ergeb,
                       headers=["Name", "B1", "B2", "B3", "B4", "G", "SP", "MP", "MP", "SP", "G", "B1", "B2", "B3",
                                "B4", "Name"]))

        # Excel-Export
        # TODO macht Programm langsam
        sht = xw.sheets(str(SpieltagNr))
        for i in range(0, len(ergeb)):
            for j in range(0, len(ergeb[i])):
                sht.range((i + 1 + zeile, j + 1)).value = ergeb[i][j]

        # Tabellenpunkte
        if ergeb[6][7] > ergeb[6][8]:
            team_a.sieg(ergeb[6][7], ergeb[6][6], ergeb[6][5])
            team_b.niederlage(ergeb[6][8], ergeb[6][9], ergeb[6][10])
        elif ergeb[6][7] == ergeb[6][8]:
            team_a.unentschieden(ergeb[6][7], ergeb[6][6], ergeb[6][5])
            team_b.unentschieden(ergeb[6][8], ergeb[6][9], ergeb[6][10])
        else:
            team_b.sieg(ergeb[6][8], ergeb[6][9], ergeb[6][10])
            team_a.niederlage(ergeb[6][7], ergeb[6][6], ergeb[6][5])

        # nicht bereite Spieler wieder anhängen
        team_a.Spieler.append(ersatz[0])
        team_a.Spieler.append(ersatz[1])
        team_b.Spieler.append(ersatz2[0])
        team_b.Spieler.append(ersatz2[1])

        return ergeb


class Table(QWidget):
    def __init__(self, Kopie):
        super().__init__()
        self.Table = QWidget
        self.left = 200
        self.top = 200
        self.width = 800
        self.height = 500
        self.Kopie = Kopie
        self.initUI(Kopie)

    def initUI(self, Kopie):
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.title = 'Tabelle'
        self.createTable(Kopie)
        # Add box layout, add table to box layout and add box layout to widget
        self.layout = QVBoxLayout()
        self.layout.addWidget(self.tableWidget)
        self.setLayout(self.layout)

        # Show widget
        self.show()

    def createTable(self, Kopie):
        # Create table
        self.tableWidget = QTableWidget()
        self.setWindowTitle(self.title)
        self.tableWidget.setRowCount(len(Kopie))
        self.tableWidget.setColumnCount(len(Kopie[1]))

        # Spaltengrößen passen sich an Platzbedarf an
        header = self.tableWidget.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        # Spalten beschriften
        self.tableWidget.setHorizontalHeaderLabels(["Name", "Punkte", "S", "U", "N", "MP", "SP", "Schnitt"])

        # Tabelle wird beschrieben
        for i in range(0, len(Kopie), 1):
            for j in range(0, len(Kopie[i]), 1):
                self.tableWidget.setItem(i, j, QTableWidgetItem(str(Kopie[i][j])))
        self.tableWidget.doubleClicked.connect(self.on_click)

    @pyqtSlot()
    def on_click(self):
        print("\n")
        for currentQTableWidgetItem in self.tableWidget.selectedItems():
            print(currentQTableWidgetItem.row(), currentQTableWidgetItem.column(), currentQTableWidgetItem.text())


class Spieltagsuebersicht(QWidget):
    def __init__(self, Kopie):
        super().__init__()
        self.Table = QWidget
        self.left = 200
        self.top = 200
        self.width = 800
        self.height = 500
        self.Kopie = Kopie
        self.initUI(Kopie)

    def initUI(self, Kopie):
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.title = 'Tabelle'
        self.createTable(Kopie)
        # Add box layout, add table to box layout and add box layout to widget
        self.layout = QVBoxLayout()
        self.layout.addWidget(self.tableWidget)
        self.setLayout(self.layout)

        # Show widget
        self.show()

    def createTable(self, Kopie):
        # Create table
        self.tableWidget = QTableWidget()
        self.setWindowTitle(self.title)
        self.tableWidget.setRowCount(len(Kopie))
        self.tableWidget.setColumnCount(len(Kopie[1]))

        # Spaltengrößen passen sich an Platzbedarf an
        header = self.tableWidget.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        # Spalten beschriften
        self.tableWidget.setHorizontalHeaderLabels(["Name", "Punkte", "S", "U", "N", "MP", "SP", "Schnitt"])

        # Tabelle wird beschrieben
        for i in range(0, len(Kopie), 1):
            for j in range(0, len(Kopie[i]), 1):
                self.tableWidget.setItem(i, j, QTableWidgetItem(str(Kopie[i][j])))
        self.tableWidget.doubleClicked.connect(self.on_click)

    @pyqtSlot()
    def on_click(self):
        print("\n")
        for currentQTableWidgetItem in self.tableWidget.selectedItems():
            print(currentQTableWidgetItem.row(), currentQTableWidgetItem.column(), currentQTableWidgetItem.text())





while 1:
    print("")
    print("1 = Saison starten")
    print("")
    inp: str = input()
    if inp == "1":
        book = xlrd.open_workbook('Input.xlsx')
        sheet = book.sheet_by_name('Tabelle1')
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
        Anzahl = 4
        Kreisoberliga = Liga(Anzahl, 1, "Kreisoberliga")
