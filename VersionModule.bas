Attribute VB_Name = "VersionModule"
' FORREST SOFTWARE
' Copyright (c) 2018 Mateusz Forrest Milewski
'
' Permission is hereby granted, free of charge,
' to any person obtaining a copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation the rights to
' use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
' and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
'
'
'
' OPIS WERSJI SEKCJA
' P.S. zapis ten byl bardzo optymistyczny i zakladal, ze wszystkim sie zajme od razu bez przerwy
' jednak smutna rzeczywistosci szybko zweryfikowala, ze bedzie spory hold na pracy na tym makrze...
' --------
' no to zaczynamy z nowym projektem, ktory ma miec pierwsza stablina wersje gotowa na czwartek (jest poniedzialek)
' bedzie to jedno z wazniejszych wyzwan jakie sobie postawie w dotychczasowej pracy na makrach.
' Bedzie to rozszerzenie makra Quarter -> 6p.
' nazwa musiala zostac zmieniona due to ilosci elementow przekazujacych informacje (z 4 na 6 sztuk)
''
' podstawowe rqm:
' zrozumiec logike dzialania poprzedniego makra - jak dane sie ukladaja - przede wszystkim ich order
' uklad wszystkich parametrow i potencjal podzielenia tych info na osobne tabele
' Z racji tego ze i tak dane te beda uzupelniane za pomoca formularzy badz od razu z innych plikow
' mozna sobie to ujednolicic
'
'
' jednym z glownych zadan kontrolnych bedzie proces pilnowania, aby wszystkie 6 elementow bylo w linii z danymi
' tak aby linki nie byly wstanie wygasnac - mysle ze bede sie opieral na parametrze znow jakiegos przyjemnego peselu


' v0 0.01
' na poczatek zdefiniowalem arkusze na ktorych bede pracowal
' zmienne globalne nazw arkuszy dostepnych w danym projekcie
' G_register_sh_nm = "register"
' G_order_release_status_sh_nm = "ORDER RELEASE STATUS"
' G_cont_pnoc_sh_nm = "Contracted . PNOC"
' G_osea_sh_nm = "OSEA"
' G_recent_build_plan_changes_sh_nm = "RECENT BUILD PLAN CHANGES"
' G_main_sh_nm = "MAIN"
' G_resp_sh_nm = "RESP"
' G_open_issues_sh_nm = "OPEN ISSUES"
' G_config_sh_nm = "config"
' G_totals_sh_nm = "TOTALS"
' G_del_conf_sh_nm = "DELIVERY CONFIRMATION"
' G_one_pager_sh_nm = "ONE PAGER"
' G_xq_sh_nm = "XQ HANDLER"


'v 0.02
' dodanie enumow przede wszystkim
' formatowanie warunkowe w arkuszu main

' v 0.03
' pierwszy fomularz glowny formmain

' v 0.04
' drugi formularz
' plus rozbudowanie eventow pod akcje zmiany danych miedzy arkuszami a formularzami

' v 0.05
' powrot po dlugim czasie
' proba zrozumienia co aktualnie posiadam na wersji 0.04...
' mala poprawka nazewnicz na subie w klasie six  p checker -> sprawdz czy aktywny to ten arkusz
' najsmieszniejsze jest to ze to powinno raczej sie nazywac to thisworkbook == activeworkbook
'
' dodatkowy formularz new edit project w arkuszu main jednak formularz reaguje za kazdy razem gdy mamy arkusz gdzie
' ejst dopasowanie pierwszych 4 kolumn - wygodna sprawa

' wciaz mam watpliwosci co do funkcjonalnosci show vbmodeless
' bardzo ograniczam tym uzytkownikow i leniwcy nie beda chcieli zaakceptowac tego pomyslu

' dodane pola w formularzu dla 1p - order release status

' prace na dzien 15 czerwca zakonczyly sie w formularzu FormOrderReleaseStatus
' na miejscu dodawania nowej linii danych - narazie prosty split i zaufanie google ze
' array zawsze zaczyna sie od indeksu zero zatem tamtejsza petla jest od zera do trzy

' doktryna jaka podjalem podczas pracy na tym makrze to przede wszystkim enkapsulacja nawet jesli logika bedzie spojna
' nie bede tworzyl zadnych pomostow
' prosta izolacja ze zlamaniem zasady DRY i pewnie odrobine KISS.

' Set o = New OrderReleaseStatusHandler
' to jeszcze nic nie daje przed uruchomieniem formularza dla order release
' wciaz nie ma zadnej logiki dzieki ktorej mozna by cokolwiek
' ruszyc co prawda w sam form mozna wpisac dane by potem one wpadly odpowiednio
' w arkusz order release jednak nie ma tam nawet zadnej konkretnej walidacji ani rowniez brak logiki
' aby poszedl feedback do arkusza main (zmiana rozowego koloru w kolumnie F, ktora ma zwierac date ostatniego update'u)

' podoba mi sie ze zielen i rozowy dzialaja na calej skali arkusza
' i rozroznia nie tylko podstawowe zalozenia ale tez jest tabularaza
' gdy nie ma danych w ogole - zatem upate z rozowego na potencjalnie zielony
' ma sie znajdowac w implementacji klasy OrderReleaseStatusHandler

' v 0.06
' 1. dodanie dwoch guzikow do ribbonu
' 2. kontynuacja implementacji:
'   a. order release status class dokonczenie implementacji dla suba wklejacego potwierdzenia do arkusza main, ze zmiany nastapily
'   b. dodanie helperow w formualrzach dzieki czemu unikamy recznego wpisywania fieldow
'
' pojawil sie problem ze spacja...
' otoz comboboxy wyboru linkow ma po przecinkach spacje... niech juz to tak zostanie
' ale powoduje ze wszelkie porownania wymagaja odpowiedniej ilosci trimow co jest nieco nuzace
' a przy okazji musze zweryfikowac jednorazowo caly kod
'
' ale tak poza tym to jestem bardzo zadowolony z tej wersji jesli chodzi o arkusz / tabele ORDER RELEASE STATUS
' wszystko dziala tak jak nalezy - kolorki dzialaja na arkusszu glownym pelna reakcja fomrularza dla tego guzika jest dobra baza
' aby teraz wszystko zmalpowac na reszte


' v 0.07
' wersja bedzie duzo kodu, ktory wlasciwie bedzie kopia kodu


' v 0.08
' handlery gotowe pod P5



' v 0.09
' 1. handlery gotowe pod P6 XQ bez testow
' 2. dodany modul i klasa czyszczenia itemow
' 3. pierwsze drafty i rusztowanie pod zaciaganie danych z aktywnych wizardow
' 4. db click na numerycznych polach formularzy
' 5. czyszczenia formularza glownego
' 6. rozwiniecie del confu
' 7. podpiecie nareszcie wszystkich guzikow dla projektu
' 8. double click dla przechodzenia miedzy arkuszami


' v 0.10
'
' pierwsze proby z synchronizacja z wizardami
' move pomiedzy arkuszami za pomoca podwojnego klikniecia
' kopiowanie danych pomiedzy rekordami


' v 0.11
'
' glownie zabawa z one pagerem zatem kodu niewiele

' v 0.12
'
' przesuniecie buttona clear item na prawo na koniec
' co by ktos przez przypadek nie nacisl
' dopisanie reszty warunkow na pozostale 6p plus resp plus open issues plus xq
' (w sumie juz nawet by trzeba liczyc 9 elementow!)
' dodanie arkuszy 3kolory oraz wizard_buff ktory w sumie jest kopia dzialania 6time

' v 0.13
'
' dodana lista country codes
' i przy okazji zostala juz ladnie wykorzytana
' w buff wizarda dodana zostala linijka rozdzielajaca country code osea

' v 0.14
' pierwsze testy logiki wkladania danych
' synchro danych z wizarda do odpwiednich pol 6p
'
' powrot do dobrego zapytania o to czy jestem pewien tego,
' ze chce zaciagnac dane z otwartego wizarda

' teraz kwestia przesypania danych z buffora wizarda bezposrednio na nowe pola kazdego arkusza


' v 0.15
' w tej wersji jeszcze to pozostanie jednak nalezy przemyslec rozdzielenie nazwy projektu od MY
' nie wiem dlaczego, ale na poczatku wydawalo mi sie to rozsadne
' jednak im glebiej wchodze w proces, tym trudniej jest ogarnac co jest co
' FMA bardzo czesto samo sobie strzela w kolano jesli chodzi o zarzadzanie danych
'
'
' kolejny topic do mozliwosc konfiguracji czy chcemy generowac raporty fma review za pomoca power pointa
' czy za pomoca kolejnych nowych plikow excelowych
' czy w ogole dac mozliwosc koordynatorom, aby sami sobie decydowali co gdzie chca zobaczyc
'
' na poczatek mysle ze fajnie by bylo miec oba, ale z drugiej strony moje lenistwo
' dazy do tego, aby dac jedno rozwiazanie kwestii
'
'w tej wersji juz pomalu klasa one pager handler jest uzupelniana 1p wlasciwie juz zaczyna dzialac jak nalezy

' zatrzylem sie na implementacji eventow podczas klikania
' na formularza wejsciowym dla one pagera musze dokonczyc eventy change'u
' na kolejnym listboxach forma pod wybieranie danych pod one pager
' to jest ostatni SUB: ListBoxPlants_Change


' v 0.16
' wersja porzadkowa
' nieco sie zmienila implmentacja jesli chodzi o form one pagera
' wyrzucilem cala logike 0.14 - byla przekombinowana


'
'
'
' v 0.17 kopia na dzien 2016-11-08
'-----------------------------------------------------------
' 0.16 wlasciwie ma gotowa logike wybierania danych pod one pager
' nie jest to moze najpiekniejsze rozwiazanie, ale caly dzien stracilem zeby jako tako sie to wszystko kleilo
' mozna filtrowac pojedynczo...
'
' krok kolejny to at last powrot do uzupelniania danych na one pagerze

' zmiana w pobieraniu danych z order release status
' jest tylko jeden rekord per faza
'-----------------------------------------------------------

'
'
' v 0.18 2016-12-13
'-----------------------------------------------------------
' dodatkowe implemntacje importu z wizarda dla kazdego formularza osobno
' duzo kodu ale za to spoko funkcjonalnosci
'-----------------------------------------------------------


' v 0.19 2016-12-19
'-----------------------------------------------------------
' baza pod zmiany na koniec stycznia 2017
' config na arkuszach configowych - dodana lista resp - reczna modyfikacja ktore FMA ma byc brane pod uwage.


' v 0.20 2016-12-22
'-----------------------------------------------------------
' kolejny milestone
' wczesniej juz powstal form resp adjuster
' jednak dopiero teraz dopisana jest jego logika
' fix na podwojnych petlach sprawdzajcych czy w wizard buff powstalo cos nowego
' wczesniej cos nie tak bylo z logika i zle przerzucalo dane
' dodatkowy test na h1 g1 - poniewaz juz jako tako nie ma miejsca w arkuszu wizard buff
' zatem wcisnalem na sama gore wartosc podliczajaca ile danych faktycznie nalezy do scopu
' jeszcze mysle jak to zastosowac do kolejnych czesci makra...
'
'
' co prawda z perspektywy layoutut pola g1 h1 nieco odstaja jesli chodzi o klarownosc, ze wydaje sie ze wszystko jest w jednym miesjcu
' ale koniec koncow nareszcie totol resp bedzie zawsze w jednym miejscu niezaleznie od ilosci danych per wizard
'
' form form resp adjuster bedzie sie pojawiac przy kazdej okazji proby importu danych z wizarda
' dla upewnienia sie czy scope ktory aktualnie zawarty jest w nim odpowiada naszym potrzebom (narazie nie wiem jak to prosciej zrobic)
' wiec bedzie to nieco toporne
'
'

'0.20 2017-01-17
'-----------------------------------------------------------
' dodanie wielu nowych elementow dla klasy one pager handler (ktory juz przekroczyl 1k linii kodu)
' pokyte 90 % zapotrzebowania na body subow zostaly pojedyncze procedury do napisania


' v 0.21 2017-01-19
'-----------------------------------------------------------
' podstawa pod kolejne rozwiniecia klasy one pager handler
' dorzucenie pozostalych implementacji dla prototypow funkcji i subow
'
' podczas testow okazalao sie ze wszystkie wykresy kolejnych wygenerowanych raportow
' odnosi sie do tego samego zrodla znajdujacego sie w raporcie macierzystym - jest to dosyc powazny
' problem poniewaz okazuje sie ze dla chronologicnzych raportow zawsze sie pokaza ostatnie dane
'
' gorzej bedzie jeszcze, jesli raporty beda sie generowaly kolejno dla roznych projektow
' okaze sie, ze seria wygenerowanych kolejnych arkuszu maja jedne wspolne wykresy ostatniego wyngenerowanego raportu
' (ostatnio projektu na liscie)


' v 0.22 2017-01-20
'-----------------------------------------------------------
' wersja 0.22 ma sprostac rozdzieleniu logiki pobierania zrodla dla wykresow
'

' v 0.23 2017-01-23
'-----------------------------------------------------------
' kopia z 0.22 z polowicznym rozwiazaniem problemu na starcie tej implementacji
' z eksperymentow manulanych wynika ze nie trzeba wiele zmieniac w source (jedynie podmienic nazwe zrodlowego arkusza)
' reszta tak jakby sama potrafila sie dopasowac (zatem licze na autonomie dzialania excela w tym przypadku)
' wersja ta zweryfikuje poprawnosc tej tezy (podstawowe testy przeszly pomyslnie, zatem narazie tak ta implementacje zrobie)
' wystarczy zmienic sam source - reszta sama powinna sie dopasowac

' v 0.24 2017-01-23
'-----------------------------------------------------------
' wersja do testu z jednym z koordynatorow

' v 0.25 2017-01-23
'-----------------------------------------------------------
' nieco lepiej rozbudowana logika resp form
' dodany guzik przerzuceania FMA resp

' v 0.26 2017-01-24
'-----------------------------------------------------------
' po pierwszych testach razem z coordami
' wstepne sprawdzenie uruchomienia dzialania
' od wersji 0.25 braki w zaciaganiu danych z wizard buffa
' pojedyncze bledy implementacji

' import w buffa wizardowego (totals) in progress
'-----------------------------------------------------------

' v 0.27 2017-01-26
'-----------------------------------------------------------
' dalsza podroz po swicie buffu wizardowego (kontynuacja totalsow)
' del conf nawet juz koloruje na czerwono jak przekrocze podejrzanie totalsy
' w sumie mozna by tak samo zrobic totale same w sobie pociagnac info z buffa z komorki h1
' sprawdzic czy total nie jest za duzy
'
' z szybkich fixow tylko napokne ze zrobilem male fo pa jesli chodzi o zmiany w formularzu wymoglem dodatkowe uruchomienie suba
' jesli total sie zmieni z czego sub wymuszal znow zmianie na totalu, czyli troche sie zapetlilem i slusznie wyrzucilo mi blad
' out of stack - musze byc nieco bardziej uwazny
'
'
' jako tako import from buffer jest prawie caly gotowy
' potrzebne sa teraz testy kordynatorow by zobaczyc czego jeszcze nie dopiescilem jesli chodzi o logike dzialania calego 6P


' v 0.28 dev 2017-01-27
'-----------------------------------------------------------
' jako pierwsza wersja w pierwszej kolejnosci idzie do testu do mnie
' jesli wszystko bedzie ok dam koordynatorom, co by mogli sie oswoic
' zrobimy szkolenie mysle pod koniec stycznia 2017 i zobaczymy jakie pytanie padna i jak samo narzedzie bedzie reagowac
'
' tez wciaz pozostaje otwarta kwestia open issues jak rowniez raportu trojkolorowego
'
' del conf import from buffer - before and after sth wrong and everything goes into open - dodana logika

' kolejny error - nie dziala guzik edycji w przypadku del confa

' v 0.29 2017-02-02
'-----------------------------------------------------------
'
' wersja kosmetyczna dopasowana pod testy kordynatorow na poczatku lutego 2017

' v 0.30 2017-02-02
'-----------------------------------------------------------
'
' wersja kosmetyczna dopasowana pod testy kordynatorow na poczatku lutego 2017

' v 0.31 2017-02-10
'-----------------------------------------------------------
'
' zmiany "usuwaniu wszystkiego"


' v 0.32 2017-02-13
'-----------------------------------------------------------
'
' wersja korekt od Marcina By
'1 w Project (kolumna A) powinno oprócz nazwy zaciagac z Wizarda czy BIW czy GA
'2. czy # of Veh. , Orders due i Released mozemy dodac do kolejnej wersji Wizarda? zaciagaloby to do 6P. Do ustalenia z wszystkimi coord.
'3. Resp Adjuster- zwiekszyc czcionke
'4. FormContractedPNOC- zaciagac dane z Wizarda - Contracted-FMA/XXX vs PNOC
'5. Totals 5P
 '- Ordered -powinna byc suma FMA, najlepiej pobierac dane z Wizarda z wpisana data przy resp FMA
 '- liczba ITDC nie pobiera sie z Wizarda
'6 '. FormDelConfStatus- status OPEN nie zlicza sie automatycznie
'7' '. NIE DA SIE PRZEJSC DO WIZARDA PRZY OTWARTYM FOMRULARZU FILL DETAILS i Project Manag.
'8'. jak liczy sie status arrived, in transit i future?


' mass import for open issues still open for development

' v 0.33 2017-02-14
'-----------------------------------------------------------
' juz w wersji 0.32 rozpoczalem implementacje dla zaciagania danych na podstawie requestow marcina b
' plus anny k - dodalem narazie prosta logike w formularzu contracted pnoc zaciagania po respie


' v 0.34 2017-02-21
'-----------------------------------------------------------
' ciagniemy dalej temat masowego zaciagania danych dla open issues
' narazie idzie to jak po grudzie bo  w sumie musze zaimplementowac calego cruda...
' w tej wersji (0.34) zrobiony delete 1 to 1


' v 0.35 2017-02-21
'-----------------------------------------------------------
' fajnie fajnie - pora na wielki powrot do mass importu z konkretnego wizarda
' jesli dobrze pamietam ustalenia klarownie okreslaly, ze zaciagamy komentarze z odowiedniej kolumny
' i robimy rekord per duns, czyli mamy sytuacje wiele do jednego proponuje zatem aby komentarze z wielu pnow
' wsadzac po dwoch enterach - potem pomyslimy jak to moze lepiej rozwiazac
' ah no i najwazniejsze ... trzeba by okreslic ktore z elementow maja byc wyswietlane
'
' mass import for open issues to test!


' v 0.36 2017-03-10
'-----------------------------------------------------------
' duzo zmian: moja logika nie byla wlasciwa...
' 1. sprawdz czy filtr na arkuszu master ma wplyw na dzialanie (tak samo z resta dla arkusza PICKUPS)
' 2. jest problem z iloscia pusow - cos kiepsko podliczyl to ostatnio w przykladzie Ewy
' 3. zle licze totale (patrze na zbyt duza ilosc danych)
' 4. open issues  ma byc zaciagane tylko z poszczegolnych rekordow - tj. dla wszystkich red, yellow (green juz nie)
' 5. del conf musi byc bardziej ogarniety - nie patrzymy globalnie tylko na tematy ktore sa w zgodzie z fma resp.

' potrzebna jest w ogole dokladniejsza logika zaciagania informacji - buffer zaciaga zbyt swobodnie


' v 0.37 2017-03-13
'-----------------------------------------------------------
' potrzebne korekty po wizycie u FMA Coordow - podliczanie z importu nie jest dokladne
' zawyzone wartosci w wielu przypadkach - w open issues nie ma byc zaciagane wszystko,
' a tylko elementy czerwone oraz zolte
' opis jest zawarty przede wszystkim na content wersji 0.36
'
'
' funkcjonalnosc mass import dla open issues jest wlasciwie gotowa - odbiera commenty tylko dla yellow i red
' dopracowanie usuwania danych z formularza open issues!
'
'-----------------------------------------------------------



' v 0.37 2017-03-13
'-----------------------------------------------------------
'
' platforma pod zamkniecie tematu bledow z poczatku 2017
' - wyrafinowana funkcjonalnosc usuwania usuwania open issues
' rozszerzenie open issues tak, aby pracowalo z danymi w zgodzie z CRUDEM
' plus mass import ogranicza sie do do zaciagania comments tylko z czesci typu yellow,
' badz red - do testow ta wersje chcemy :)
'
'-----------------------------------------------------------


' v 0.38 2017-03-21
'-----------------------------------------------------------
''
'
' kolejna odslona zabezpieczajaca - open issues - prawie gotowe
'
'
' still some issues: V0.39
' --------------------------------------------------------------------------------------


' --------------------------------------------------------------------------------------
'
'' V 0.40 - NOWE DEL CONFY -  zgodnie ze zmianami w collectorze i wizardzie
'-----------------------------------------------------------

' --------------------------------------------------------------------------------------
'
'' V 0.41 - small graph changes
'-----------------------------------------------------------


' --------------------------------------------------------------------------------------
'
'' V 0.42 - fixy po kolejnym spotkaniu z coords

' - zmiana globalna projektu - by auto dopasowac reszte arkuszy
' - MY nie moze byc z przecinkiem - lepiej za wczasu blokowac takie rzeczy
' - problem z przerzucaniem danych trojkolorwych - jest jakies dziwne przesuniecie danych
' - TOTAL zle policzony
' - open issues - brak formatowania dla wiekszej ilosci textu
'-----------------------------------------------------------
'

' --------------------------------------------------------------------------------------
'' V 0.43 -

' - kolejna faza rozbudowy mass importu dla wszystkich arkusz w jednym strzale
' --------------------------------------------------------------------------------------

' --------------------------------------------------------------------------------------
'' V 0.44

' powersja prototypu rozszerzona o mozliwosc zaciagania info z starej wersji raportu (Q)
' --------------------------------------------------------------------------------------

' --------------------------------------------------------------------------------------

' --------------------------------------------------------------------------------------
' v 0.45
' wersja z udanym elementem migracyjnym
' --------------------------------------------------------------------------------------
' v 0.46
' fix na one pagerze - zle dane wstawia
' - szczegolnie na czesci recent build plan changes
' - i del confach
' --------------------------------------------------------------------------------------


' --------------------------------------------------------------------------------------
' v 0.47
'
' fix na del conf na one pager

' --------------------------------------------------------------------------------------


' --------------------------------------------------------------------------------------
' v 0.48
'
' - fix on checkbox for adding new and do mass import - sth wrong with shiet is...
' - proba wygenerowana pustych danych generuje petle nieskonczona
' - po importcie msgbox by sie przydal
' - do testu mass import dla open issues - zmiany ukladu kolumn w wizardzie? - done?
' - freeze panes added

' --------------------------------------------------------------------------------------



' --------------------------------------------------------------------------------------
' v 0.49
'
' - mass import should work now
' - fix na nieskonczonej petli

' --------------------------------------------------------------------------------------

' --------------------------------------------------------------------------------------
' v 0.5
'
' - uproszczenie schematu mass import: open issues - brak agresji w doborze nazwy projektu

' --------------------------------------------------------------------------------------


' --------------------------------------------------------------------------------------
' v 0.51
'
' - dla roznych faz zle tworzy One Pagera...

' --------------------------------------------------------------------------------------



' --------------------------------------------------------------------------------------
' v 0.52
'
' - dopasowanie mismatchy dla configa kasi
' - wpisanie foo dla sortowania po phase list jednak jej ostateczne nie-uzycie - poniewaz
' - najpierw dane trzeba w kolekcji poukladac chronologicznie

' --------------------------------------------------------------------------------------


' --------------------------------------------------------------------------------------
' v 0.53
'
' Po pierwszym fma review 2018-02-28:
' - zle zlicza sie total
' - poprawa logiki dla pus recv, it, future
' - stestowac sorotwanie pierwszej tabeli (order release status)
' - poprawic widocznosc komentarzy dla open issues - wiecej miejsca?
' - zrobic open issues jako element opcjonalny
' - przygotowac sie do zmiany sekcji del conf
' - opisac nazwa po prostu del conf
' - z tego co widze to logika osea nie jest do konca zrobiona
' - dodtkowa funkcjonalnosc lean - ktora usuwa puste zaciagniecia z quartera
' - kopiowanie danych pomiedzy 6p
' - power point nie pobiera headera
' --------------------------------------------------------------------------------------


' --------------------------------------------------------------------------------------
' v 0.54
'
' Po pierwszym fma review 2018-03-05
' - zle zlicza sie total - OK
' - poprawa logiki dla pus recv, it, future - NOK
' - stestowac sorotwanie pierwszej tabeli (order release status) - OK
' - poprawic widocznosc komentarzy dla open issues - wiecej miejsca? - OK
' - zrobic open issues jako element opcjonalny - NOK
' - przygotowac sie do zmiany sekcji del conf - NOK
' - opisac nazwa po prostu del conf - OK
' - z tego co widze to logika osea nie jest do konca zrobiona - OK
' - dodtkowa funkcjonalnosc lean - ktora usuwa puste zaciagniecia z quartera - OK
' - kopiowanie danych pomiedzy 6p - OK
' - power point nie pobiera headera - OK
' --------------------------------------------------------------------------------------
