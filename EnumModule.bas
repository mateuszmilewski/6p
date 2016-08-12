Attribute VB_Name = "EnumModule"
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


' glowne ENUMy pracujace w tym projekcie
' ciekawostka jest wspolny enum dla pierwszych 4 wystapien
' jest to traktowane jako link *wykorzystujemy obiekt T_Link

Public Enum E_LINK_ORDER
    e_link_project = 1
    e_link_plt = 2
    e_link_faza = 3
    e_link_cw = 4
End Enum


Public Enum E_MAIN_ORDER
    e_main_status = 5
    e_main_last_update_on_order_release_status
    e_main_last_update_on_recent_build_plan_changes
    e_main_last_update_on_chart_contracted_pnoc
    e_main_last_update_on_osea
    e_main_last_update_on_totals
    e_main_last_update_on_xq
    e_main_last_update_on_del_conf
    e_main_last_update_on_open_issues
    e_main_last_update_on_resp
End Enum


Public Enum E_ORDER_RELEASE_STATUS
    e_order_release_mrd = 5
    e_order_release_build
    e_order_release_bom_freeze
    e_order_release_no_of_veh
    e_order_release_orders_due
    e_order_release_released
    e_order_release_weeks_delay
End Enum

Public Enum E_RECENT_BP_CH
    e_recent_bp_ch_no_of_veh = 5
    e_recent_bp_ch_tbw
    e_recent_bp_ch_order_release_ch
    e_recent_bp_ch_comment
End Enum

Public Enum E_CONT_PNOC_CHART
    e_cont_pnoc_chart_contracted = 5
    e_cont_pnoc_chart_pnoc
    e_cont_pnoc_chart_open_bp
    e_cont_pnoc_chart_actionable_fma
End Enum

Public Enum E_OSEA_ORDER
    e_osea_order_total = 5
    e_osea_order_on_stock
    e_osea_order_ordered
    e_osea_order_confirmed
    e_osea_order_for_mrd
    e_osea_order_after_mrd
    e_osea_order_open
End Enum

Public Enum e_5p_totals
    e_5p_total = 5
    e_5p_na
    e_5p_itdc
    e_5p_pnoc
    e_5p_fma_eur
    e_5p_fma_osea
    e_5p_ordered
    e_5p_arrived
    e_5p_in_transit
    e_5p_ppap_status
    e_5p_no_ppap_status
End Enum


Public Enum E_XQ_ORDER
    e_xq_comment = 5
    e_xq_ppap_gate
    e_xq_project_type
End Enum


Public Enum E_DEL_CONF_ORDER


    e_del_conf_on_stock = 5
    e_del_conf_edi
    e_del_conf_ho
    e_del_conf_na
    

    e_del_conf_for_mrd
    e_del_conf_after_mrd
    
    e_del_conf_for_smrd
    e_del_conf_after_smrd
    
    e_del_conf_for_twomrd
    e_del_conf_after_twomrd
    
    e_del_conf_for_twosmrd
    e_del_conf_after_twosmrd
    
    e_del_conf_for_alt
    e_del_conf_after_alt
    
    e_del_conf_open
    e_del_conf_pot_itdc
    e_del_conf_undef
    
    
    
    
End Enum

Public Enum E_OPEN_ISSUES_ORDER
    e_open_issues_status = 5
    e_open_issues_no_of_pn
    e_open_issues_part_supplier
    e_open_issues_delivery
    e_open_issues_comment
End Enum

Public Enum E_RESP_ORDER
    e_resp_fma = 5
    e_resp_osea
    e_resp_pem
    e_resp_ppm
    e_resp_sqe
End Enum
