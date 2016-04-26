Attribute VB_Name = "VersionModule"
' FORREST SOFTWARE
' Copyright (c) 2015 Mateusz Forrest Milewski
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

