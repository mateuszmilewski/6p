Attribute VB_Name = "GlobalModule"
' FORREST SOFTWARE
' Copyright (c) 2016 Mateusz Forrest Milewski
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

' zmienne globalne nazw arkuszy dostepnych w danym projekcie
Global Const G_register_sh_nm = "register"
Global Const G_order_release_status_sh_nm = "ORDER RELEASE STATUS"
Global Const G_cont_pnoc_sh_nm = "Contracted . PNOC"
Global Const G_osea_sh_nm = "OSEA"
Global Const G_recent_build_plan_changes_sh_nm = "RECENT BUILD PLAN CHANGES"
Global Const G_main_sh_nm = "MAIN"
Global Const G_resp_sh_nm = "RESP"
Global Const G_open_issues_sh_nm = "OPEN ISSUES"
Global Const G_config_sh_nm = "config"
Global Const G_totals_sh_nm = "TOTALS"
Global Const G_del_conf_sh_nm = "DELIVERY CONFIRMATION STATUS"
Global Const G_one_pager_sh_nm = "ONE PAGER"
' ----------------------------------------------------------
' ----------------------------------------------------------
Global Const G_chart1_handler_sh_nm = "CHART1_Handler"
Global Const G_chart2_handler_sh_nm = "CHART2_Handler"
Global Const G_chart3_handler_sh_nm = "CHART3_Handler"

' nazwy chartow raportu
' ----------------------------------------------------------
Global Const G_CHART_1_NM = "Chart PNOC"
Global Const G_CHART_2_NM = "Chart OSEA"
Global Const G_CHART_3_NM = "Chart Total"
' ----------------------------------------------------------

' ----------------------------------------------------------
' ----------------------------------------------------------
Global Const G_xq_sh_nm = "XQ HANDLER"

Global Const G_PHASE_LIST_SH_NM = "PHASE_LIST"
Global Const G_WIZARD_BUFF_SH_NM = "WIZARD_BUFF"
Global Const G_DEL_CONF_SPECIAL_SH_NM = "delivery_confirmation_special"

Global Const G_CC_SH_NM = "CountryCodes"


Global Const POLOWA_CAPACITY_ARKUSZA = 524288 ' 2^19 polowa capacity akursza excela
Global Const CAPACITY_ARKUSZA = 1048576


' text na guziki
Global Const G_BTN_TEXT_ADD = "Dodaj"
Global Const G_BTN_TEXT_EDIT = "Edytuj"

Global Const G_ORDER_RELEASE_STATUS_LIMIT = 8
Global Const G_RECENT_BUILD_PLAN_CHANGES_LIMIT = 4
Global Const G_CHART1_CONTRACTED_PNOC_LIMIT = 5
Global Const G_OPEN_ISSUES_LIMIT = 23



Global Const G_SEPARATOR = "qwertyuiopasdfghjkl"


Global Const G_MAIN_TB_FROM_Q = "main_table"






