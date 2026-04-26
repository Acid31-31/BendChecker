# BendChecker

WPF/.NET Tool zur Analyse von Blech-STEP-Modellen (Kanten/Machbarkeit).

## Voraussetzungen
- Visual Studio 2022 (aktuell)
- .NET 8 SDK

## Setup
1. Repo klonen
2. `BendChecker.sln` in Visual Studio öffnen
3. NuGet Packages wiederherstellen
4. Build & Run

## Status
V1: UI + Excel-Regeln + Regel-Auswahl + Findings-Liste.

## Nächste Schritte
- STEP/OCCT-Analyse integrieren (CascadeSharp): Biegezonen (Zylinderflächen) erkennen, Flanschlängen messen
- Werkzeugprüfung (Standard 90° Matrize, 88° Stempel, r=1mm) und später eigene Werkzeuge per JSON/Excel