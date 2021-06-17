# RobotikkMCServer
En MC verden med server for robotikkstudenter

## Hvordan bruke serveren:
1. Pull siste versjon fra Git
2. Kjør ´´´ Run.bat ´´´ eller ´´´ Standard.bat ´´´. 
    * Se førstegangsinstallasjon 
3. Da skal det komme et kommandovindu opp uten noen feilmeldinger 
4. For å teste om den fungerer kan du logge inn med ip-en: ´´´ localhost ´´´
5. For å stoppe serveren skriv: ´´´ stop ´´´ i kommandovinduet.

## Førstegangsinstallasjon:
1. Finn filen ´´´ Standard.bat ´´´ i repoet.
2. Kopier den og kall kopien ´´´ Run.bat ´´´ (Dette er for at din personlige fil skal påvirke andre)
3. I filen kan du endre ´´´ -Xmx{tall}{prefix} ´´´ og ´´´ -Xms{tall}{prefix} ´´´. Dette er hvor mye ram serveren har tilgang til (jo mer jo bedre for spillerne på serveren men værre for deg)
    * Xmx er maks ram den kan bruke
    * Xms er hvor mye ram den kan satarte med??
    * Prefiksene kan være G (giga) eller MB (mega)
