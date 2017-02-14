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

' heading
' to jest remark kopii czesci klasy one pager handler
' Set pcs(1) = ThisWorkbook.Sheets(SIXP.G_main_sh_nm)
' left side
' Set pcs(2) = ThisWorkbook.Sheets(SIXP.G_order_release_status_sh_nm)
' Set pcs(3) = ThisWorkbook.Sheets(SIXP.G_cont_pnoc_sh_nm)
' Set pcs(4) = ThisWorkbook.Sheets(SIXP.G_osea_sh_nm)
' right side
' Set pcs(5) = ThisWorkbook.Sheets(SIXP.G_recent_build_plan_changes_sh_nm)
' Set pcs(6) = ThisWorkbook.Sheets(SIXP.G_totals_sh_nm)
' Set pcs(7) = ThisWorkbook.Sheets(SIXP.G_del_conf_sh_nm)
' bottom
' Set pcs(8) = ThisWorkbook.Sheets(SIXP.G_resp_sh_nm)
' second page
' Set pcs(9) = ThisWorkbook.Sheets(SIXP.G_open_issues_sh_nm)



Public Enum E_PPAP
    
    E_PPAP_NOK = 0
    E_PPAP_OK = 1
    
End Enum


Public Enum E_SPECIAL_CASE_FOR_DEL_CONF
    E_SPEC_CASE_TAKE_DEL_CONF_CONNECTED_WITH_MRD = 10
    E_SPEC_CASE_DO_NOT_TAKE_DEL_CONF_CONNECTED_WITH_MRD = 100
    E_SPEC_CASE_COUNT_BEFORE_AND_AFTER_MRD = 1000
End Enum

Public Enum E_MATCH
    E_LIKE
    E_EQUAL
    E_BEFORE_OR_AFTER_MRD
End Enum

Public Enum E_PCS_ORDER
    e_pcs_main = 1
    e_pcs_order_release_status
    e_pcs_pnoc
    e_pcs_osea
    e_pcs_recent_build_plan_changes
    e_pcs_totals
    e_pcs_delivery_confirmation
    e_pcs_resp
    e_pcs_open_issues
End Enum

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
    e_5p_future
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
    e_del_conf_too_late
    e_del_conf_pot_itdc
    
    
    
    
    
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







' wizard high order
Public Enum E_MASTER_MANDATORY_COLUMNS
    pn = 1
    Alternative_PN
    PN_Name
    GPDS_PN_Name
    duns
    Supplier_Name
    country_code
    MGO_code
    Responsibility
    fup_code
    SQ
    ppap_status
    SQ_Comments
    MRD1_QTY
    MRD2_QTY
    Total_QTY
    ADD_to_T_slash_D
    MRD1_Ordered_date
    MRD1_Ordered_QTY
    MRD1_Ordered_STATUS
    MRD1_confirmed_qty
    MRD1_confirmed_qty_dot__Status
    MRD1_Total_PUS_STATUS
    MRD2_Ordered_date
    MRD2_Ordered_QTY
    MRD2_Ordered_STATUS
    MRD2_confirmed_qty
    MRD2_confirmed_qty_dot__Status
    MRD2_Total_PUS_STATUS
    Delivery_confirmation
    First_Confirmed_PUS_Date
    Delivery_reconfirmation
    Total_PUS_QTY
    Total_PUS_STATUS
    Comments
    Bottleneck
    Future_Osea
    DRE
    EDI_Received
    Capacity
    Oncost_confirmation
    BLANK3
    BLANK4
End Enum

Public Enum E_NEW_PROJECT_ITEM
    plt = 1
    project
    biw_ga ' BIW or GA
    MY
    PHAZE
    BOM
    PICKUP_DATE
    PPAP_GATE
    mrd
    BUILD_START
    BUILD_END
    KOORDYNATOR
    E_ACTIVE
    CAPACITY_CHECK
    E_MRD_DATE
    E_MRD_REG_ROUTES
    E_PLATFORM
    E_TRANSPORTATION_ACCOUNT_NUMBER
    E_UNIQUE_ID
End Enum


Public Enum E_PUS_SH
    O_INDX = 1
    O_PN
    O_DUNS
    O_FUP_code
    O_Pick_up_date
    O_Delivery_Date
    O_Pick_up_Qty
    O_PUS_Number
End Enum


Public Enum E_ONE_PAGERS_INTO
    E_ONE_PAGERS_INTO_POWER_POINT
    E_ONE_PAGERS_INTO_SEPERATE_EXCELS
End Enum

Public Enum E_SORT_TYPE
    E_ASCENDING
    E_DESCENDING
End Enum

Public Enum E_DEL_CONF_SPECIAL
    E_DCS_BLANK = 2
    E_DCS_POTITDC
    E_DCS_MRD
    E_DCS_Staggered_MRD
    E_DCS_TWO_Staggered_MRD
    E_DCS_TWO_MRD
    E_DCS_HO
    E_DCS_EDI
    E_DCS_ON_STOCK
    E_DCS_NA
    E_DCS_ALT_MRD
    E_DCS_OPEN
    E_DCS_OPEN_TOO_LATE

End Enum
