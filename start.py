import random
import sys
from copy import deepcopy
from typing import List, Any

import PyQt5.QtGui
import matplotlib.pyplot as plt
import numpy as np
import xlrd
import xlwings as xw
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QApplication, QWidget, QTableWidget, QTableWidgetItem, QVBoxLayout, QHeaderView, QPushButton
from tabulate import tabulate


# TODO Sven Kettenbeil       147   136   122   147   552     2     0     1     2   551   138   153   124   136  Oliver Heinold
# TODO Normalverteilung anhand von ALLEN Einzelergebnissen einer Liga berechnen


# TODO Conditional Formatting einbauen für Gesamtergebnis

# TODO Talent einbauen in Alterung, Stärkeanpassung einschätzen zwecks Realismus
# TODO Stärkeänderung bissel abschwächen
# TODO Peak-Alter als Zufallswert einbauen & an Normalverteilung übergebn
# TODO Neugeneration Spieler, Gute Saison (Schnitt > Stärke) -> mehr Verbesserung und umgekehrt


# TODO langsamer Anzeigemodus / Bahn für Bahn / Starter für Starter


# TODO Chance erhöhen wenn in Stammformation (erste6), Chance erhöhen wenn in alten Team NICHT in Startformation
# vllt. nur 1 Versuch pro Spieltag??
# TODO überprüfen Annäherung Schnitt und Stärke nach 26 Spieltagen, dann einschätzen wie dringend Form implementiert werden muss
# TODO Geld nach Saisonerhalt und gewisse Schwelle damit Spieler wechselt?
# TODO Geld als zusätzlicher Anreiz, zum Chance erhöhen??

# TODO Ergebnisse als GUI mit farblichen Ergebnissen
# TODO GUI Doppelklick auf Spieler mit Graph von Ergebnissen
# TODO Schnittliste Liga??
# TODO mehrere Klassen Promotion / relegation (als erstes noch über manuelle Eingabe / später automatisch)
# TODO Formkurve Spieler, alle paar Woche mal schlechte Form z.B.
# TODO Saisoncounter, jede Saison neues Tabellenblatt mit Rostern und Tabellen
# TODO Bei Spieler-Abruf ALLE Ergebnisse saisonübergreifend darstellen + Graph
# TODO verstecktes Talent, beurteilen anhand wie Graph sich verändert hat.
# TODO Team managen und Austellung bestimmen
# TODO Bahnrekorde, Bestwerte der Spieler speichern
# TODO Heimbahn Wählitz zb besser als Teuchern
# TODO Pokal
# TODO Geld für Transfers
# TODO Neurales Netzwerk lernt was es machen muss um erster in Tabelle zu werden, vorher umstellen Spielerwechsel über Name zu Spielerwechsel nach i???


# TODO Bei Wählitz Gesamtschnitte///Heimschnitte als Stärke und sehen wie sie damit aufsteigen, wie weit sie kommen???


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
                    Array = [data[i + j][3], data[i + j][4], data[i + j][5], data[i + j][6], 0, 0, 0]
                    for k in range(8, sheet.ncols, 1):
                        Array.append(data[i + j][k])
                    spieler.append(Array)
                team_a = Verein(data[i][1], data[i][2], spieler, 0, 0, 0, 0, 0, 0)
                self.Ligaa.append(team_a)
        spielplan = self.spielplanGenerator(anzahl)

        self.menu(spielplan, spieltag_nr)
        self.alterung()

        # aus Objekten zu Excel
        try:
            wb = xw.Book("Input.xlsx")
            sht = xw.sheets("Tabelle1")
            print("Moment,speichern.....")
            for i in range(0, len(self.Ligaa) * 8, 8):
                sht.range((i + 1, 1)).value = self.Liganame
                sht.range((i + 1, 2)).value = self.Ligaa[int(i / 8)].Name
                for j in range(0, 8, 1):
                    sht.range((i + j + 1, 4)).value = self.Ligaa[int(i / 8)].Spieler[j]
            wb.save()  # speichern, da sonst Änderungen nicht geladen werden
            wb.close()
        except:
            print("Bitte Input-Datei schließen, und mit 0 bestätigen")
            eing = input()
            wb = xw.Book("Input.xlsx")
            sht = xw.sheets("Tabelle1")
            print("Moment,speichern.....")
            for i in range(0, len(self.Ligaa) * 8, 8):
                sht.range((i + 1, 1)).value = self.Liganame
                sht.range((i + 1, 2)).value = self.Ligaa[int(i / 8)].Name
                for j in range(0, 8, 1):
                    sht.range((i + j + 1, 4)).value = self.Ligaa[int(i / 8)].Spieler[j]
            wb.save()  # speichern, da sonst Änderungen nicht geladen werden
            wb.close()

    def menu(self, spielplan, spieltag_nr):
        # Team für langsamen Modus
        print("Zu überwachendes Team: ")
        Slow = input()
        while 1:
            print("")
            print("1 = Team hinzufügen")
            print("2 = Statistik Spieler")
            print("3 = Statistik Team")
            print("4 = Tabelle")
            print("5 = Spieltag")
            print("50 = Beobachtetes Team ändern")
            print("6 = Spielerwechsel")
            print("60 = Spielerwechsel Godmode")
            print("0 = Saison beenden und Teams speichern")
            print("")
            inp: str = input()
            if inp == "1":
                self.Teamadd()
            elif inp == "2":
                self.statistikSpieler()
            elif inp == "3":
                self.statistikTeam()
            elif inp == "4":
                self.Tabelle(spieltag_nr)
            elif inp == "5":
                vorbei = self.Spieltag(spielplan, spieltag_nr, Slow)
                if vorbei == 2:
                    spieltag_nr -= 1
                spieltag_nr += 1
            elif inp == "50":
                Slow = self.Slow()
            elif inp == "6":
                self.Spielerwechsel()
            elif inp == "60":
                self.SpielerwechselGodmode()
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

    def statistikSpieler(self):
        print("Spieler: ")
        Sp1 = input()
        i1 = -1

        # Position im Array finden
        # für jeden Verein
        for i in range(0, len(self.Ligaa), 1):
            # für jeden Spieler des Vereins
            for j in range(0, len(self.Ligaa[i].Spieler), 1):
                # wenn Name richtig
                if (self.Ligaa[i].Spieler[j][0] == Sp1):
                    Array = deepcopy(self.Ligaa[i].Spieler[j])
                    # Nicht-Ergebnisse entfernen
                    del Array[0:7]
                    # leere Zellen entfernen
                    while ("" in Array):
                        Array.remove("")
                    print(Array)
                    plt.plot(Array)
                    plt.ylabel('Ergebnisse')
                    plt.show()
                    i1 = 0
                    break
        if (i1 == -1):
            print("Spieler nicht gefunden")
            return 0

    # TeamStatistik ausgeben
    def statistikTeam(self):
        print("Welcher Verein?")
        inp = input()
        for Verein in self.Ligaa:
            if Verein.Name == inp:
                print(Verein.Name)
                # deepcopy, da sonst Originalarray verändert wurde
                Array = deepcopy(Verein.Spieler)
                # entfernen der Spalte für "Stärke" und Talent
                for i in range(0, len(Array)):
                    del Array[i][1]
                    del Array[i][2]
                print(tabulate(Array))
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

    def Spieltag(self, Spielplan, SpieltagNr, Slow):
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
                # try:
                if (Slow == self.Ligaa[Spielplan[SpieltagNr - 1][i][0] - 1].Name or Slow == self.Ligaa[
                    Spielplan[SpieltagNr - 1][i][1] - 1].Name):
                    app = QApplication(sys.argv)
                    Spiel2 = SpielSlow(self.Ligaa[Spielplan[SpieltagNr - 1][i][0] - 1],
                                       self.Ligaa[Spielplan[SpieltagNr - 1][i][1] - 1], zeile, SpieltagNr)
                    # Wichtige Zeile, damit Fenster offen bleibt, aber nach Schließung das Programm weiterläuft
                    app.exec_()
                else:
                    Spiel2 = Spiel(self.Ligaa[Spielplan[SpieltagNr - 1][i][0] - 1],
                                   self.Ligaa[Spielplan[SpieltagNr - 1][i][1] - 1], zeile, SpieltagNr)
                ergeb = Spiel2.Ausgabe()
                print(tabulate(ergeb,
                               headers=["Name", "B1", "B2", "B3", "B4", "G", "SP", "MP", "MP", "SP", "G", "B1",
                                        "B2", "B3",
                                        "B4", "Name"]))
                zeile += 10
                wb.save()
            # except:
            #    print("Anzahl Teams Spielplan und Anzahl Teams Liga stimmen nicht überein")
            #   return

    def Slow(self):
        print("Zu überwachendes Team: ")
        self.Slow = input()
        return self.Slow


    def Spielerwechsel(self):

        Wechsel = 0.3

        print("wird versucht zu verpflichten: ")
        Sp1 = input()
        i1 = -1
        j1 = 0
        StärkeT1 = 0
        # Position im Array Spieler 1 finden
        # für jeden Verein
        for i in range(0, len(self.Ligaa), 1):
            # für jeden Spieler des Vereins
            for j in range(0, len(self.Ligaa[i].Spieler), 1):
                # wenn Name richtig
                if (self.Ligaa[i].Spieler[j][0] == Sp1):
                    i1 = i
                    j1 = j
                    # Stärke des Teams des zu verpflichtenden Spielers ermitteln
                    # nicht in erster Spielerschleife, da sonst ALLE Vereine addiert weden
                    for k in range(0, len(self.Ligaa[i].Spieler), 1):
                        StärkeT1 += self.Ligaa[i].Spieler[k][1]
                    break
        if (i1 == -1 and j1 == 0):
            print("Spieler nicht gefunden")
            return 0

        print("soll abgegeben werden: ")
        Sp2 = input()
        i2 = -1
        j2 = 0
        StärkeT2 = 0
        # Position im Array Spieler 1 finden
        for i in range(0, len(self.Ligaa), 1):
            for j in range(0, len(self.Ligaa[i].Spieler), 1):
                if (self.Ligaa[i].Spieler[j][0] == Sp2):
                    i2 = i
                    j2 = j
                    for k in range(0, len(self.Ligaa[i].Spieler), 1):
                        StärkeT2 += self.Ligaa[i].Spieler[k][1]
                    break
        if (i2 == -1 and j2 == 0):
            print("Spieler nicht gefunden")
            return 0

        # Wenn Teams das verpflichten will, stärker ist als altes Team, höhere Chance
        if (StärkeT2 >= StärkeT1):
            Wechsel += ((StärkeT2 / StärkeT1) - 1) * 5
            print(str(round(((StärkeT2 / StärkeT1) - 1) * 500)) + "% dadurch, dass neues Team besser ist als altes")
        else:
            Wechsel += ((StärkeT2 / StärkeT1) - 1) * 10
            print(str(round(((StärkeT2 / StärkeT1) - 1) * 1000)) + "% dadurch, dass altes Team besser ist als neues")

        # Wenn Spieler schlechter ist als andere Spieler im neuen Team
        if (StärkeT2 >= self.Ligaa[i1].Spieler[j1][1] * 8):
            Wechsel += ((StärkeT2 / (self.Ligaa[i1].Spieler[j1][1] * 8)) - 1) * 3
            print(str(round(((StärkeT2 / (self.Ligaa[i1].Spieler[j1][
                                              1] * 8)) - 1) * 300)) + "% dadurch, dass Spieler schlechter ist als andere im neuem Team")
        else:
            Wechsel += ((StärkeT2 / (self.Ligaa[i1].Spieler[j1][1] * 8)) - 1) * 3
            print(str(round(((StärkeT2 / (self.Ligaa[i1].Spieler[j1][
                                              1] * 8)) - 1) * 300)) + "% dadurch, dass Spieler besser ist als andere im neuem Team")

        print(str(round(Wechsel * 100)) + "% Wechselchance")

        rand = random.randint(0, 100)
        if (rand <= Wechsel * 100):
            Wechsel = 1
        print(rand)

        if Wechsel == 1:
            self.Ligaa[i1].Spieler.append(self.Ligaa[i2].Spieler[j2])
            self.Ligaa[i2].Spieler.append(self.Ligaa[i1].Spieler[j1])
            self.Ligaa[i1].Spieler.pop(j1)
            self.Ligaa[i2].Spieler.pop(j2)
            print("Spielerwechsel erfolgreich!")
        else:
            print("Spieler wechselt nicht!")

    def SpielerwechselGodmode(self):

        print("wird versucht zu verpflichten: ")
        Sp1 = input()
        i1 = -1
        j1 = 0
        # Position im Array Spieler 1 finden
        # für jeden Verein
        for i in range(0, len(self.Ligaa), 1):
            # für jeden Spieler des Vereins
            for j in range(0, len(self.Ligaa[i].Spieler), 1):
                # wenn Name richtig
                if (self.Ligaa[i].Spieler[j][0] == Sp1):
                    i1 = i
                    j1 = j
                    break
        if (i1 == -1 and j1 == 0):
            print("Spieler nicht gefunden")
            return 0

        print("soll abgegeben werden: ")
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
        print("Spielerwechsel erfolgreich!")

    def alterung(self):
        for i in range(0, len(self.Ligaa)):
            print(" ")
            print(self.Ligaa[i].Name)
            print(" ")
            for j in range(0, len(self.Ligaa[i].Spieler)):
                self.Ligaa[i].Spieler[j][2] += 1
                erg = np.random.normal((28 - self.Ligaa[i].Spieler[j][2]) / 2, 1, 1) / 100 * self.Ligaa[i].Spieler[j][
                    1]
                neu = erg + self.Ligaa[i].Spieler[j][1]

                print(self.Ligaa[i].Spieler[j][0] + " " + str(round(self.Ligaa[i].Spieler[j][2])) + " alt: " + str(
                    round(self.Ligaa[i].Spieler[j][1])) + " Verä.: " + str(round(erg[0])) + " neu: " + str(
                    round(neu[0])))

                # Neue Stärke eintragen
                self.Ligaa[i].Spieler[j][1] = neu[0]


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
        self.ergeb = [[(0) for c in range(0, 16)] for r in range(0, 7)]

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

        for j in range(0, 6, 1):
            # Namen und Alter eintragen
            self.ergeb[j][0] = "(" + str(int(team_a.Spieler[j][2])) + ")" + str(team_a.Spieler[j][0])
            self.ergeb[j][15] = "(" + str(int(team_b.Spieler[j][2])) + ")" + str(team_b.Spieler[j][0])
            # Tagesform, beeinflusst Gesamtergebnis Spieler
            hrand_a = np.random.normal(1000, 30, 1)
            grand_a = np.random.normal(1000, 30, 1)
            for i in range(1, 5, 1):
                hrand = np.random.normal(1000, 70, 1)
                self.ergeb[j][i] = int(team_a.Spieler[j][1] * (hrand_a / 1000) * (hrand / 1000) / 4)
                # Ergebnis Spieler
                self.ergeb[j][5] += self.ergeb[j][i]

                grand = np.random.normal(1000, 70, 1)
                self.ergeb[j][i + 10] = int(team_b.Spieler[j][1] * (grand_a / 1000) * (grand / 1000) / 4)
                # Ergebnis Spieler
                self.ergeb[j][10] += self.ergeb[j][i + 10]
                # SP Spieler
                if self.ergeb[j][i + 10] > self.ergeb[j][i]:
                    self.ergeb[j][9] += 1
                elif self.ergeb[j][i + 10] == self.ergeb[j][i]:
                    self.ergeb[j][9] += 0.5
                # MP Spieler
                if self.ergeb[j][9] > 2:
                    self.ergeb[j][8] = 1
                elif self.ergeb[j][9] == 2:
                    if self.ergeb[j][10] > self.ergeb[j][5]:
                        self.ergeb[j][8] = 1
                    elif self.ergeb[j][10] == self.ergeb[j][5]:
                        self.ergeb[j][8] = 0.5

            # für statistikTeam, Anzahl Spiele und Gesamtholz hochzählen, Schnitt berechnen
            team_a.Spieler[j][4] += self.ergeb[j][5]
            team_a.Spieler[j][5] += 1
            team_a.Spieler[j][6] = team_a.Spieler[j][4] / team_a.Spieler[j][5]
            team_a.Spieler[j].append(self.ergeb[j][5])
            team_b.Spieler[j][4] += self.ergeb[j][10]
            team_b.Spieler[j][5] += 1
            team_b.Spieler[j][6] = team_b.Spieler[j][4] / team_b.Spieler[j][5]
            team_b.Spieler[j].append(self.ergeb[j][10])

            # Punkte Heim
            # SP Spieler
            self.ergeb[j][6] = 4 - self.ergeb[j][9]
            # MP Spieler
            self.ergeb[j][7] = 1 - self.ergeb[j][8]
            # Gesamtholz Mannschaft
            self.ergeb[6][5] += self.ergeb[j][5]
            self.ergeb[6][10] += self.ergeb[j][10]
            # MP Gesamt
            self.ergeb[6][7] += self.ergeb[j][7]
            self.ergeb[6][8] += self.ergeb[j][8]
            # SP Gesa,t
            self.ergeb[6][6] += self.ergeb[j][6]
            self.ergeb[6][9] += self.ergeb[j][9]
        # MP für Gesamtholz
        if self.ergeb[6][5] > self.ergeb[6][10]:
            self.ergeb[6][7] += 2
        elif self.ergeb[6][5] == self.ergeb[6][10]:
            self.ergeb[6][7] += 1
            self.ergeb[6][8] += 1
        else:
            self.ergeb[6][8] += 2

        # Excel-Export

        try:
            sht = xw.sheets(str(SpieltagNr))
            # TODO macht Programm langsam
            for i in range(0, len(self.ergeb)):
                for j in range(0, len(self.ergeb[i])):
                    sht.range((i + 1 + zeile, j + 1)).value = self.ergeb[i][j]
        except:
            print("Bitte Input-Datei schließen!")

        # Tabellenpunkte
        if self.ergeb[6][7] > self.ergeb[6][8]:
            team_a.sieg(self.ergeb[6][7], self.ergeb[6][6], self.ergeb[6][5])
            team_b.niederlage(self.ergeb[6][8], self.ergeb[6][9], self.ergeb[6][10])
        elif self.ergeb[6][7] == self.ergeb[6][8]:
            team_a.unentschieden(self.ergeb[6][7], self.ergeb[6][6], self.ergeb[6][5])
            team_b.unentschieden(self.ergeb[6][8], self.ergeb[6][9], self.ergeb[6][10])
        else:
            team_b.sieg(self.ergeb[6][8], self.ergeb[6][9], self.ergeb[6][10])
            team_a.niederlage(self.ergeb[6][7], self.ergeb[6][6], self.ergeb[6][5])

        # nicht bereite Spieler wieder anhängen
        team_a.Spieler.append(ersatz[0])
        team_a.Spieler.append(ersatz[1])
        team_b.Spieler.append(ersatz2[0])
        team_b.Spieler.append(ersatz2[1])

    def Ausgabe(self):
        return self.ergeb


class SpielSlow(QWidget):
    def __init__(self, team_a, team_b, zeile, SpieltagNr):
        super().__init__()
        self.Table = QWidget
        self.left = 200
        self.top = 200
        self.width = 800
        self.height = 500
        self.TeamA = team_a
        self.TeamB = team_b
        self.zeile = zeile
        self.SpieltagNr = SpieltagNr
        self.i = 0
        self.j = 1
        self.Heim = 0
        self.Gast = 0

        self.ergeb = Spiel(self.TeamA, self.TeamB, zeile, SpieltagNr).Ausgabe()

        self.initUI()

    def initUI(self):
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.title = 'Spiel'
        self.createTable()
        # Add box layout, add table to box layout and add box layout to widget
        self.layout = QVBoxLayout()
        self.layout.addWidget(self.tableWidget)
        self.setLayout(self.layout)
        btn = QPushButton("weiter", self)
        btn.move(50, 400)
        btn.clicked.connect(self.weiter)

        # Show widget
        self.show()

    def createTable(self):
        # Create table
        self.tableWidget = QTableWidget()
        self.setWindowTitle(self.title)
        self.tableWidget.setRowCount(7)
        self.tableWidget.setColumnCount(16)

        # Spaltengrößen passen sich an Platzbedarf an
        header = self.tableWidget.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        for i in range(0, 16):
            header.setSectionResizeMode(i, QHeaderView.ResizeToContents)
        # Spalten beschriften
        self.tableWidget.setHorizontalHeaderLabels(
            ["Name", "B1", "B2", "B3", "B4", "G", "SP", "MP", "MP", "SP", "G", "B1", "B2", "B3",
             "B4", "Name"])

        # Tabelle wird beschrieben

        for i in range(0, len(self.ergeb), 1):
            self.tableWidget.setItem(i, 0, QTableWidgetItem(str(self.ergeb[i][0])))
            self.tableWidget.setItem(i, 15, QTableWidgetItem(str(self.ergeb[i][15])))

    def weiter(self):
        if self.i > 4:
            pass
        else:
            # Heim
            self.tableWidget.setItem(self.i, self.j, QTableWidgetItem(str(self.ergeb[self.i][self.j])))
            self.tableWidget.setItem(self.i + 1, self.j, QTableWidgetItem(str(self.ergeb[self.i + 1][self.j])))
            self.Heim += self.ergeb[self.i][self.j] + self.ergeb[self.i + 1][self.j]

            # Gast
            self.tableWidget.setItem(self.i, self.j + 10, QTableWidgetItem(str(self.ergeb[self.i][self.j + 10])))
            self.tableWidget.setItem(self.i + 1, self.j + 10,
                                     QTableWidgetItem(str(self.ergeb[self.i + 1][self.j + 10])))
            self.Gast += self.ergeb[self.i][self.j + 10] + self.ergeb[self.i + 1][self.j + 10]

            COLORS = ['#053061', '#2166ac', '#4393c3', '#92c5de', '#d1e5f0', '#f7f7f7', '#fddbc7', '#f4a582', '#d6604d',
                      '#b2182b', '#67001f']

            if (self.ergeb[self.i][self.j] > self.ergeb[self.i][self.j + 10]):
                self.tableWidget.item(self.i, self.j).setBackground(PyQt5.QtGui.QColor(255, 250, 205))
            elif (self.ergeb[self.i][self.j] > self.ergeb[self.i][self.j + 10]):
                self.tableWidget.item(self.i, self.j).setBackground(PyQt5.QtGui.QColor(240, 255, 255))
                self.tableWidget.item(self.i, self.j + 10).setBackground(PyQt5.QtGui.QColor(240, 255, 255))
            else:
                self.tableWidget.item(self.i, self.j + 10).setBackground(PyQt5.QtGui.QColor(255, 250, 205))

            if (self.ergeb[self.i + 1][self.j] > self.ergeb[self.i + 1][self.j + 10]):
                self.tableWidget.item(self.i + 1, self.j).setBackground(PyQt5.QtGui.QColor(255, 250, 205))
            elif (self.ergeb[self.i + 1][self.j] > self.ergeb[self.i + 1][self.j + 10]):
                self.tableWidget.item(self.i + 1, self.j).setBackground(PyQt5.QtGui.QColor(240, 255, 255))
                self.tableWidget.item(self.i + 1, self.j + 10).setBackground(PyQt5.QtGui.QColor(240, 255, 255))
            else:
                self.tableWidget.item(self.i + 1, self.j + 10).setBackground(PyQt5.QtGui.QColor(255, 250, 205))

            # Gesamtholz Heim und Gast
            self.tableWidget.setItem(6, 6, QTableWidgetItem(str(self.Heim)))
            self.tableWidget.setItem(6, 9, QTableWidgetItem(str(self.Gast)))
            self.tableWidget.setItem(6, 5, QTableWidgetItem(str(self.Heim - self.Gast)))

            # Wenn alle Bahnen gespielt, Punkte anzeigen
            if self.j == 4:
                for k in range(1, 7):
                    # Heim
                    self.tableWidget.setItem(self.i, self.j + k, QTableWidgetItem(str(self.ergeb[self.i][self.j + k])))
                    self.tableWidget.setItem(self.i + 1, self.j + k,
                                             QTableWidgetItem(str(self.ergeb[self.i + 1][self.j + k])))
                self.i += 2
                self.j = 1
            else:
                self.j += 1

    def Ausgabe(self):
        return self.ergeb


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
