Option Strict Off
Option Explicit On

Public Class clsNSK0000H_sql
    '/----------------------------------------------------------------------/"
    '/"
    '/      クラス     ：SQL関連NSK0000Hクラス"
    '/      ＩＤ       ：clsNSK0000H_sql"
    '/      概要       ：SQL文に関する関数群"
    '/"
    '/      作成者： CHRISTOPHER CREATE 2017/04/27       REV 01.00" 
    '/      更新者： CHRISTOPHER UPDATE 2017/05/02       REV 01.01"
    '/      更新者： L. Mapula   UPDATE 2017/05/11 
    '/      更新者： Richard     UPDATE 2017/05/24 
    '/                        更新内容：( Updated parameter names and datatypes )"
    '/"
    '/     Copyright (C) Inter co.,ltd 1997"
    '/----------------------------------------------------------------------/"
    Public w_sqlBuilder As System.Text.StringBuilder
    Public w_Sql As String
    '2018/02/28 Yamanshi Del -----------------------------------------
    'Private Enum gKinmuTimeMCol As Integer
    '    WorkTimeNum        '実勤務時間
    '    DayTime            '日勤時間
    '    NightTime          '夜勤時間
    '    NextNightTime      '翌日夜勤時間
    '    TotalNightTime     '総夜勤時間
    '    NextTotalNightTime '翌夜勤総時間
    '    FirstFromTime      '前半開始時刻
    '    SecToTime          '後半終了時刻
    'End Enum
    '-----------------------------------------------------------------
    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <param name="p_intPlanNo"></param>
    ''' <param name="p_strStaffDataID"></param>
    ''' <returns></returns>
    Public Function select_NS_STAFFHISTORY_F_02(ByRef w_Rs As ADODB.Recordset,
                                                ByVal p_intPlanNo As Integer,
                                                ByVal p_strStaffDataID As String) As Boolean
        select_NS_STAFFHISTORY_F_02 = False

        Try
            w_sqlBuilder = New System.Text.StringBuilder

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "SELECT   SKILLLVLCD "
                w_Sql = w_Sql & ",DISPNO1 "
                w_Sql = w_Sql & ",DISPNO2 "
                w_Sql = w_Sql & ",DISPNO3 "
                w_Sql = w_Sql & ",DISPNO4 "
                w_Sql = w_Sql & ",DISPNO5 "
                w_Sql = w_Sql & ",AUTOALLOCKBN "
                w_Sql = w_Sql & ",TEAM "
                w_Sql = w_Sql & ",NIGHTONLYSTAFFKBN "
                w_Sql = w_Sql & ",PATTERNCD "
                w_Sql = w_Sql & "FROM NS_STAFFHISTORY_F "
                w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "AND PLANNO = " & p_intPlanNo & " "
                w_Sql = w_Sql & "AND KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "' "
                w_Sql = w_Sql & "AND STAFFMNGID = '" & p_strStaffDataID & "' "
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "SELECT   SKILLLVLCD "
                w_Sql = w_Sql & ",DISPNO1 "
                w_Sql = w_Sql & ",DISPNO2 "
                w_Sql = w_Sql & ",DISPNO3 "
                w_Sql = w_Sql & ",DISPNO4 "
                w_Sql = w_Sql & ",DISPNO5 "
                w_Sql = w_Sql & ",AUTOALLOCKBN "
                w_Sql = w_Sql & ",TEAM "
                w_Sql = w_Sql & ",NIGHTONLYSTAFFKBN "
                w_Sql = w_Sql & ",PATTERNCD "
                w_Sql = w_Sql & "FROM NS_STAFFHISTORY_F "
                w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "AND PLANNO = " & p_intPlanNo & " "
                w_Sql = w_Sql & "AND KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "' "
                w_Sql = w_Sql & "AND STAFFMNGID = '" & p_strStaffDataID & "' "
            End If

            w_Rs = General.paDBRecordSetOpen(w_Sql)
            select_NS_STAFFHISTORY_F_02 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <param name="p_strGiryoCD"></param>
    ''' <returns></returns>
    Public Function select_NS_SKILL_M_01(ByRef w_Rs As ADODB.Recordset,
                                         ByVal p_strGiryoCD As String) As Boolean
        select_NS_SKILL_M_01 = False

        Try
            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "Select DispNo"
                w_Sql = w_Sql & ", SkillBunruiCD"
                w_Sql = w_Sql & " From NS_SKILL_M"
                w_Sql = w_Sql & " Where SkillLvlCD = '" & p_strGiryoCD & "'"
                w_Sql = w_Sql & " And HospitalCD = '" & General.g_strHospitalCD & "'"
                w_Sql = w_Sql & " "
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "Select DispNo"
                w_Sql = w_Sql & ", SkillBunruiCD"
                w_Sql = w_Sql & " From NS_SKILL_M"
                w_Sql = w_Sql & " Where SkillLvlCD = '" & p_strGiryoCD & "'"
                w_Sql = w_Sql & " And HospitalCD = '" & General.g_strHospitalCD & "'"
                w_Sql = w_Sql & " "
            End If

            w_Rs = General.paDBRecordSetOpen(w_Sql)
            select_NS_SKILL_M_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <returns></returns>
    Public Function select_NS_NECESSARYNUM_M_01(ByRef w_Rs As ADODB.Recordset) As Boolean
        select_NS_NECESSARYNUM_M_01 = False
        Try
            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "SELECT TEAM"
                w_Sql = w_Sql & ", KINMUCD"
                w_Sql = w_Sql & ", MONNINZU"
                w_Sql = w_Sql & ", TUENINZU"
                w_Sql = w_Sql & ", WEDNINZU"
                w_Sql = w_Sql & ", THUNINZU"
                w_Sql = w_Sql & ", FRININZU"
                w_Sql = w_Sql & ", SATNINZU"
                w_Sql = w_Sql & ", SUNNINZU"
                w_Sql = w_Sql & ", HOLNINZU"
                w_Sql = w_Sql & " FROM NS_NECESSARYNUM_M"
                w_Sql = w_Sql & " WHERE KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "'"
                w_Sql = w_Sql & " AND HOSPITALCD = '" & General.g_strHospitalCD & "'"
                w_Sql = w_Sql & " ORDER BY TEAM, KINMUCD"
                w_Sql = w_Sql & " "
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "SELECT TEAM"
                w_Sql = w_Sql & ", KINMUCD"
                w_Sql = w_Sql & ", MONNINZU"
                w_Sql = w_Sql & ", TUENINZU"
                w_Sql = w_Sql & ", WEDNINZU"
                w_Sql = w_Sql & ", THUNINZU"
                w_Sql = w_Sql & ", FRININZU"
                w_Sql = w_Sql & ", SATNINZU"
                w_Sql = w_Sql & ", SUNNINZU"
                w_Sql = w_Sql & ", HOLNINZU"
                w_Sql = w_Sql & " FROM NS_NECESSARYNUM_M"
                w_Sql = w_Sql & " WHERE KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "'"
                w_Sql = w_Sql & " AND HOSPITALCD = '" & General.g_strHospitalCD & "'"
                w_Sql = w_Sql & " ORDER BY TEAM, KINMUCD"
                w_Sql = w_Sql & " "
            End If

            w_Rs = General.paDBRecordSetOpen(w_Sql)
            select_NS_NECESSARYNUM_M_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <param name="p_intPlanNo"></param>
    ''' <returns></returns>
    Public Function select_NS_PERSONALKINMUCOND_F_01(ByRef w_Rs As ADODB.Recordset,
                                                     ByVal p_intPlanNo As Integer) As Boolean
        select_NS_PERSONALKINMUCOND_F_01 = False

        Try

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "SELECT "
                w_Sql = w_Sql & " PKC.STAFFMNGID, "
                w_Sql = w_Sql & " PKC.KINMUCD, "
                w_Sql = w_Sql & " PKC.ALLOCCHKFLG, "
                w_Sql = w_Sql & " PKC.COUNTMAX, "
                w_Sql = w_Sql & " PKC.COUNTMIN, "
                w_Sql = w_Sql & " PKC.MONSPECIFY, "
                w_Sql = w_Sql & " PKC.TUESPECIFY, "
                w_Sql = w_Sql & " PKC.WEDSPECIFY, "
                w_Sql = w_Sql & " PKC.THUSPECIFY, "
                w_Sql = w_Sql & " PKC.FRISPECIFY, "
                w_Sql = w_Sql & " PKC.SATSPECIFY, "
                w_Sql = w_Sql & " PKC.SUNSPECIFY, "
                w_Sql = w_Sql & " PKC.CONTINUECOUNTMAX "
                w_Sql = w_Sql & "FROM NS_PERSONALKINMUCOND_F PKC "
                w_Sql = w_Sql & "INNER JOIN NS_STAFFHISTORY_F SH ON "
                w_Sql = w_Sql & " SH.HOSPITALCD = PKC.HOSPITALCD "
                w_Sql = w_Sql & " AND SH.PLANNO = " & p_intPlanNo & " "
                w_Sql = w_Sql & " AND SH.KINMUDEPTCD = PKC.KINMUDEPTCD "
                w_Sql = w_Sql & " AND SH.STAFFMNGID = PKC.STAFFMNGID "
                w_Sql = w_Sql & " AND SH.AUTOALLOCKBN = '1' "
                w_Sql = w_Sql & " AND SH.PATTERNCD IS NULL "
                w_Sql = w_Sql & "WHERE "
                w_Sql = w_Sql & " PKC.KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "' "
                w_Sql = w_Sql & " AND PKC.HOSPITALCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "ORDER BY PKC.STAFFMNGID, PKC.KINMUCD "
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "SELECT "
                w_Sql = w_Sql & " PKC.STAFFMNGID, "
                w_Sql = w_Sql & " PKC.KINMUCD, "
                w_Sql = w_Sql & " PKC.ALLOCCHKFLG, "
                w_Sql = w_Sql & " PKC.COUNTMAX, "
                w_Sql = w_Sql & " PKC.COUNTMIN, "
                w_Sql = w_Sql & " PKC.MONSPECIFY, "
                w_Sql = w_Sql & " PKC.TUESPECIFY, "
                w_Sql = w_Sql & " PKC.WEDSPECIFY, "
                w_Sql = w_Sql & " PKC.THUSPECIFY, "
                w_Sql = w_Sql & " PKC.FRISPECIFY, "
                w_Sql = w_Sql & " PKC.SATSPECIFY, "
                w_Sql = w_Sql & " PKC.SUNSPECIFY, "
                w_Sql = w_Sql & " PKC.CONTINUECOUNTMAX "
                w_Sql = w_Sql & "FROM NS_PERSONALKINMUCOND_F PKC "
                w_Sql = w_Sql & "INNER JOIN NS_STAFFHISTORY_F SH ON "
                w_Sql = w_Sql & " SH.HOSPITALCD = PKC.HOSPITALCD "
                w_Sql = w_Sql & " AND SH.PLANNO = " & p_intPlanNo & " "
                w_Sql = w_Sql & " AND SH.KINMUDEPTCD = PKC.KINMUDEPTCD "
                w_Sql = w_Sql & " AND SH.STAFFMNGID = PKC.STAFFMNGID "
                w_Sql = w_Sql & " AND SH.AUTOALLOCKBN = '1' "
                w_Sql = w_Sql & " AND SH.PATTERNCD IS NULL "
                w_Sql = w_Sql & "WHERE "
                w_Sql = w_Sql & " PKC.KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "' "
                w_Sql = w_Sql & " AND PKC.HOSPITALCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "ORDER BY PKC.STAFFMNGID, PKC.KINMUCD "
            End If

            w_Rs = General.paDBRecordSetOpen(w_Sql)
            select_NS_PERSONALKINMUCOND_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <returns></returns>
    Public Function select_NS_CONDITION_F_03(ByRef w_Rs As ADODB.Recordset) As Boolean
        select_NS_CONDITION_F_03 = False

        Try
            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "SELECT CONDCD"
                w_Sql = w_Sql & " FROM NS_CONDITION_F"
                w_Sql = w_Sql & " WHERE APPLYKBN = '1'"
                w_Sql = w_Sql & " AND KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "'"
                w_Sql = w_Sql & " AND HOSPITALCD = '" & General.g_strHospitalCD & "'"
                w_Sql = w_Sql & " "
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "SELECT CONDCD"
                w_Sql = w_Sql & " FROM NS_CONDITION_F"
                w_Sql = w_Sql & " WHERE APPLYKBN = '1'"
                w_Sql = w_Sql & " AND KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "'"
                w_Sql = w_Sql & " AND HOSPITALCD = '" & General.g_strHospitalCD & "'"
                w_Sql = w_Sql & " "
            End If

            w_Rs = General.paDBRecordSetOpen(w_Sql)
            select_NS_CONDITION_F_03 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <returns></returns>
    Public Function select_NS_SKILLSET_F_01(ByRef w_Rs As ADODB.Recordset) As Boolean
        select_NS_SKILLSET_F_01 = False

        Try
            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "SELECT NS_SKILLSET_F.KINMUCD, NS_SKILLSET_F.SEQ"
                w_Sql = w_Sql & ", NS_SKILLSET_F.SKILLNINZU1"
                w_Sql = w_Sql & ", NS_SKILLSET_F.SKILLNINZU2"
                w_Sql = w_Sql & ", NS_SKILLSET_F.SKILLNINZU3"
                w_Sql = w_Sql & ", NS_SKILLSET_F.SKILLNINZU4"
                w_Sql = w_Sql & ", NS_KINMUNAME_M.MARKF"
                w_Sql = w_Sql & " FROM NS_SKILLSET_F, NS_CONDITION_F, NS_KINMUNAME_M"
                w_Sql = w_Sql & " WHERE NS_SKILLSET_F.CONDCD = NS_CONDITION_F.CONDCD"
                w_Sql = w_Sql & " AND NS_SKILLSET_F.KINMUDEPTCD = NS_CONDITION_F.KINMUDEPTCD"
                w_Sql = w_Sql & " AND NS_SKILLSET_F.HOSPITALCD = NS_CONDITION_F.HOSPITALCD"
                w_Sql = w_Sql & " AND NS_KINMUNAME_M.HOSPITALCD = NS_CONDITION_F.HOSPITALCD"
                w_Sql = w_Sql & " AND NS_KINMUNAME_M.KINMUCD = NS_SKILLSET_F.KINMUCD"
                w_Sql = w_Sql & " AND NS_CONDITION_F.APPLYKBN = '1'"
                w_Sql = w_Sql & " AND NS_CONDITION_F.KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "'"
                w_Sql = w_Sql & " AND NS_CONDITION_F.HOSPITALCD = '" & General.g_strHospitalCD & "'"
                w_Sql = w_Sql & " ORDER BY NS_SKILLSET_F.KINMUCD,NS_SKILLSET_F.SEQ"
                w_Sql = w_Sql & " "
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "SELECT NS_SKILLSET_F.KINMUCD, NS_SKILLSET_F.SEQ"
                w_Sql = w_Sql & ", NS_SKILLSET_F.SKILLNINZU1"
                w_Sql = w_Sql & ", NS_SKILLSET_F.SKILLNINZU2"
                w_Sql = w_Sql & ", NS_SKILLSET_F.SKILLNINZU3"
                w_Sql = w_Sql & ", NS_SKILLSET_F.SKILLNINZU4"
                w_Sql = w_Sql & ", NS_KINMUNAME_M.MARKF"
                w_Sql = w_Sql & " FROM NS_SKILLSET_F, NS_CONDITION_F, NS_KINMUNAME_M"
                w_Sql = w_Sql & " WHERE NS_SKILLSET_F.CONDCD = NS_CONDITION_F.CONDCD"
                w_Sql = w_Sql & " AND NS_SKILLSET_F.KINMUDEPTCD = NS_CONDITION_F.KINMUDEPTCD"
                w_Sql = w_Sql & " AND NS_SKILLSET_F.HOSPITALCD = NS_CONDITION_F.HOSPITALCD"
                w_Sql = w_Sql & " AND NS_KINMUNAME_M.HOSPITALCD = NS_CONDITION_F.HOSPITALCD"
                w_Sql = w_Sql & " AND NS_KINMUNAME_M.KINMUCD = NS_SKILLSET_F.KINMUCD"
                w_Sql = w_Sql & " AND NS_CONDITION_F.APPLYKBN = '1'"
                w_Sql = w_Sql & " AND NS_CONDITION_F.KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "'"
                w_Sql = w_Sql & " AND NS_CONDITION_F.HOSPITALCD = '" & General.g_strHospitalCD & "'"
                w_Sql = w_Sql & " ORDER BY NS_SKILLSET_F.KINMUCD,NS_SKILLSET_F.SEQ"
                w_Sql = w_Sql & " "
            End If
            w_Rs = General.paDBRecordSetOpen(w_Sql)
            select_NS_SKILLSET_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    '2018/02/27 Yamanishi Upd Start ------------------------------------------------------------------------------------------------------
    '''' <summary>
    '''' NSK0000HA
    '''' </summary>
    '''' <param name="w_Rs"></param>
    '''' <param name="p_intTo"></param>
    '''' <param name="p_intPlanNo"></param>
    '''' <param name="p_intStartDate"></param>
    '''' <returns></returns>
    'Public Function select_NS_STAFFBASISINFO_F_01(ByRef w_Rs As ADODB.Recordset,
    '                                              ByVal p_intTo As Integer,
    '                                              ByVal p_intPlanNo As Integer,
    '                                              ByVal p_intStartDate As Integer) As Boolean
    '    select_NS_STAFFBASISINFO_F_01 = False

    '    Try
    '        w_sqlBuilder = New System.Text.StringBuilder

    '        With w_sqlBuilder
    '            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
    '                .AppendLine("SELECT SB.STAFFMNGID")
    '                .AppendLine("     , SB.STAFFNAME")
    '                .AppendLine("     , KP.DATEF")
    '                .AppendLine("  FROM NS_STAFFBASISINFO_F SB")
    '                .AppendLine("    INNER JOIN NS_PLANCONTROL_F PC")
    '                .AppendLine("        ON  PC.HOSPITALCD = SB.HOSPITALCD")
    '                .AppendLine("    INNER JOIN NS_KINMUPLAN_F KP")
    '                .AppendLine("        ON  KP.HOSPITALCD = SB.HOSPITALCD")
    '                .AppendLine("        AND KP.STAFFMNGID = SB.STAFFMNGID")
    '                .AppendLine("        AND KP.DATEF  >= PC.PLANPERIODFROM")
    '                .AppendLine("        AND (KP.DATEF <= PC.PLANPERIODTO")
    '                .AppendLine("          OR NOT EXISTS(SELECT * FROM NS_KINMURESULT_F KR")
    '                .AppendLine("                         WHERE KR.HOSPITALCD = SB.HOSPITALCD")
    '                .AppendLine("                           AND KR.STAFFMNGID = SB.STAFFMNGID")
    '                .AppendLine("                           AND KP.DATEF BETWEEN PC.PLANPERIODTO + 1")
    '                .AppendLine("                                            AND " & p_intTo & "))")
    '                .AppendLine(" WHERE SB.HOSPITALCD = '" & General.g_strHospitalCD & "'")
    '                .AppendLine("   AND PC.PLANNO     =  " & p_intPlanNo)
    '                .AppendLine("   AND KP.OUENKINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "'")
    '                .AppendLine("   AND KP.DATEF     BETWEEN  " & p_intStartDate)
    '                .AppendLine("                        AND  " & p_intTo)
    '                .AppendLine("UNION ALL")
    '                .AppendLine("SELECT SB.STAFFMNGID")
    '                .AppendLine("     , SB.STAFFNAME")
    '                .AppendLine("     , KR.DATEF")
    '                .AppendLine("  FROM NS_STAFFBASISINFO_F SB")
    '                .AppendLine("    INNER JOIN NS_PLANCONTROL_F PC")
    '                .AppendLine("        ON  PC.HOSPITALCD = SB.HOSPITALCD")
    '                .AppendLine("    INNER JOIN NS_KINMURESULT_F KR")
    '                .AppendLine("        ON  KR.HOSPITALCD = SB.HOSPITALCD")
    '                .AppendLine("        AND KR.STAFFMNGID = SB.STAFFMNGID")
    '                .AppendLine("        AND NOT KR.DATEF BETWEEN PC.PLANPERIODFROM")
    '                .AppendLine("                             AND PC.PLANPERIODTO")
    '                .AppendLine(" WHERE SB.HOSPITALCD = '" & General.g_strHospitalCD & "'")
    '                .AppendLine("   AND PC.PLANNO     =  " & p_intPlanNo)
    '                .AppendLine("   AND KR.OUENKINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "'")
    '                .AppendLine("   AND KR.DATEF     BETWEEN  " & p_intStartDate)
    '                .AppendLine("                        AND  " & p_intTo)
    '                .AppendLine(" ORDER BY STAFFMNGID")
    '            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
    '                .AppendLine("SELECT SB.STAFFMNGID")
    '                .AppendLine("     , SB.STAFFNAME")
    '                .AppendLine("     , KP.DATEF")
    '                .AppendLine("  FROM NS_STAFFBASISINFO_F SB")
    '                .AppendLine("    INNER JOIN NS_PLANCONTROL_F PC")
    '                .AppendLine("        ON  PC.HOSPITALCD = SB.HOSPITALCD")
    '                .AppendLine("    INNER JOIN NS_KINMUPLAN_F KP")
    '                .AppendLine("        ON  KP.HOSPITALCD = SB.HOSPITALCD")
    '                .AppendLine("        AND KP.STAFFMNGID = SB.STAFFMNGID")
    '                .AppendLine("        AND KP.DATEF  >= PC.PLANPERIODFROM")
    '                .AppendLine("        AND (KP.DATEF <= PC.PLANPERIODTO")
    '                .AppendLine("          OR NOT EXISTS(SELECT * FROM NS_KINMURESULT_F KR")
    '                .AppendLine("                         WHERE KR.HOSPITALCD = SB.HOSPITALCD")
    '                .AppendLine("                           AND KR.STAFFMNGID = SB.STAFFMNGID")
    '                .AppendLine("                           AND KP.DATEF BETWEEN PC.PLANPERIODTO + 1")
    '                .AppendLine("                                            AND " & p_intTo & "))")
    '                .AppendLine(" WHERE SB.HOSPITALCD = '" & General.g_strHospitalCD & "'")
    '                .AppendLine("   AND PC.PLANNO     =  " & p_intPlanNo)
    '                .AppendLine("   AND KP.OUENKINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "'")
    '                .AppendLine("   AND KP.DATEF     BETWEEN  " & p_intStartDate)
    '                .AppendLine("                        AND  " & p_intTo)
    '                .AppendLine("UNION ALL")
    '                .AppendLine("SELECT SB.STAFFMNGID")
    '                .AppendLine("     , SB.STAFFNAME")
    '                .AppendLine("     , KR.DATEF")
    '                .AppendLine("  FROM NS_STAFFBASISINFO_F SB")
    '                .AppendLine("    INNER JOIN NS_PLANCONTROL_F PC")
    '                .AppendLine("        ON  PC.HOSPITALCD = SB.HOSPITALCD")
    '                .AppendLine("    INNER JOIN NS_KINMURESULT_F KR")
    '                .AppendLine("        ON  KR.HOSPITALCD = SB.HOSPITALCD")
    '                .AppendLine("        AND KR.STAFFMNGID = SB.STAFFMNGID")
    '                .AppendLine("        AND NOT KR.DATEF BETWEEN PC.PLANPERIODFROM")
    '                .AppendLine("                             AND PC.PLANPERIODTO")
    '                .AppendLine(" WHERE SB.HOSPITALCD = '" & General.g_strHospitalCD & "'")
    '                .AppendLine("   AND PC.PLANNO     =  " & p_intPlanNo)
    '                .AppendLine("   AND KR.OUENKINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "'")
    '                .AppendLine("   AND KR.DATEF     BETWEEN  " & p_intStartDate)
    '                .AppendLine("                        AND  " & p_intTo)
    '                .AppendLine(" ORDER BY STAFFMNGID")
    '            End If
    '        End With
    '        w_Sql = w_sqlBuilder.ToString
    '        w_Rs = General.paDBRecordSetOpen(w_Sql)
    '        select_NS_STAFFBASISINFO_F_01 = True
    '    Catch ex As Exception
    '        Throw
    '    End Try
    'End Function

    '''' <summary>
    '''' NSK0000HA
    '''' </summary>
    '''' <param name="w_Rs"></param>
    '''' <param name="p_intFrom"></param>
    '''' <param name="p_intTo"></param>
    '''' <returns></returns>
    'Public Function select_NS_STAFFBASISINFO_F_02(ByRef w_Rs As ADODB.Recordset,
    '                                              ByVal p_intFrom As Integer,
    '                                              ByVal p_intTo As Integer) As Boolean
    '    select_NS_STAFFBASISINFO_F_02 = False

    '    Try
    '        w_sqlBuilder = New System.Text.StringBuilder
    '        With w_sqlBuilder

    '            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
    '                .AppendLine("SELECT SB.STAFFMNGID")
    '                .AppendLine("     , SB.STAFFNAME")
    '                .AppendLine("     , KR.DATEF")
    '                .AppendLine("  FROM NS_STAFFBASISINFO_F SB")
    '                .AppendLine("    INNER JOIN NS_KINMURESULT_F KR")
    '                .AppendLine("        ON  KR.STAFFMNGID = SB.STAFFMNGID")
    '                .AppendLine("        AND KR.HOSPITALCD = SB.HOSPITALCD")
    '                .AppendLine(" WHERE SB.HOSPITALCD      = '" & General.g_strHospitalCD & "'")
    '                .AppendLine("   AND KR.OUENKINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "'")
    '                .AppendLine("   AND KR.DATEF     BETWEEN  " & p_intFrom)
    '                .AppendLine("                        AND  " & p_intTo)
    '                .AppendLine(" ORDER BY SB.STAFFMNGID")
    '            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
    '                .AppendLine("SELECT SB.STAFFMNGID")
    '                .AppendLine("     , SB.STAFFNAME")
    '                .AppendLine("     , KR.DATEF")
    '                .AppendLine("  FROM NS_STAFFBASISINFO_F SB")
    '                .AppendLine("    INNER JOIN NS_KINMURESULT_F KR")
    '                .AppendLine("        ON  KR.STAFFMNGID = SB.STAFFMNGID")
    '                .AppendLine("        AND KR.HOSPITALCD = SB.HOSPITALCD")
    '                .AppendLine(" WHERE SB.HOSPITALCD      = '" & General.g_strHospitalCD & "'")
    '                .AppendLine("   AND KR.OUENKINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "'")
    '                .AppendLine("   AND KR.DATEF     BETWEEN  " & p_intFrom)
    '                .AppendLine("                        AND  " & p_intTo)
    '                .AppendLine(" ORDER BY SB.STAFFMNGID")
    '            End If
    '        End With
    '        w_Sql = w_sqlBuilder.ToString
    '        w_Rs = General.paDBRecordSetOpen(w_Sql)
    '        select_NS_STAFFBASISINFO_F_02 = True
    '    Catch ex As Exception
    '        Throw
    '    End Try
    'End Function

    '''' <summary>
    '''' NSK0000HA
    '''' </summary>
    '''' <param name="w_Rs"></param>
    '''' <param name="p_blnSearchKakuteiData"></param>
    '''' <param name="p_strTargetTable"></param>
    '''' <param name="p_intTo"></param>
    '''' <param name="p_intFrom"></param>
    '''' <returns></returns>
    'Public Function select_NS_STAFFBASISINFO_F_03(ByRef w_Rs As ADODB.Recordset,
    '                                              ByVal p_blnSearchKakuteiData As Boolean,
    '                                              ByVal p_strTargetTable As String,
    '                                              ByVal p_intTo As Integer,
    '                                              ByVal p_intFrom As Integer) As Boolean
    '    select_NS_STAFFBASISINFO_F_03 = False

    '    Try

    '        If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
    '            w_Sql = "Select NS_STAFFBASISINFO_F.STAFFMNGID"
    '            w_Sql = w_Sql & ", NS_STAFFBASISINFO_F.STAFFNAME"
    '            If p_blnSearchKakuteiData = False Then
    '                p_strTargetTable = "NS_KINMUPLAN_F"
    '            Else
    '                p_strTargetTable = "NS_KINMURESULT_F"
    '            End If
    '            w_Sql = w_Sql & ", " & p_strTargetTable & ".DATEF"
    '            w_Sql = w_Sql & " FROM NS_STAFFBASISINFO_F, " & p_strTargetTable
    '            w_Sql = w_Sql & " WHERE " & p_strTargetTable & ".OUENKINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "'"
    '            w_Sql = w_Sql & " AND " & p_strTargetTable & ".DATEF <= " & p_intTo
    '            w_Sql = w_Sql & " AND " & p_strTargetTable & ".DATEF >= " & p_intFrom
    '            w_Sql = w_Sql & " AND " & p_strTargetTable & ".STAFFMNGID = NS_STAFFBASISINFO_F.STAFFMNGID"
    '            w_Sql = w_Sql & " AND " & p_strTargetTable & ".HOSPITALCD = NS_STAFFBASISINFO_F.HOSPITALCD"
    '            w_Sql = w_Sql & " AND NS_STAFFBASISINFO_F.HOSPITALCD = '" & General.g_strHospitalCD & "'"
    '            w_Sql = w_Sql & " ORDER BY NS_STAFFBASISINFO_F.STAFFMNGID"
    '        ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
    '            w_Sql = "Select NS_STAFFBASISINFO_F.STAFFMNGID"
    '            w_Sql = w_Sql & ", NS_STAFFBASISINFO_F.STAFFNAME"
    '            If p_blnSearchKakuteiData = False Then
    '                p_strTargetTable = "NS_KINMUPLAN_F"
    '            Else
    '                p_strTargetTable = "NS_KINMURESULT_F"
    '            End If
    '            w_Sql = w_Sql & ", " & p_strTargetTable & ".DATEF"
    '            w_Sql = w_Sql & " FROM NS_STAFFBASISINFO_F, " & p_strTargetTable
    '            w_Sql = w_Sql & " WHERE " & p_strTargetTable & ".OUENKINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "'"
    '            w_Sql = w_Sql & " AND " & p_strTargetTable & ".DATEF <= " & p_intTo
    '            w_Sql = w_Sql & " AND " & p_strTargetTable & ".DATEF >= " & p_intFrom
    '            w_Sql = w_Sql & " AND " & p_strTargetTable & ".STAFFMNGID = NS_STAFFBASISINFO_F.STAFFMNGID"
    '            w_Sql = w_Sql & " AND " & p_strTargetTable & ".HOSPITALCD = NS_STAFFBASISINFO_F.HOSPITALCD"
    '            w_Sql = w_Sql & " AND NS_STAFFBASISINFO_F.HOSPITALCD = '" & General.g_strHospitalCD & "'"
    '            w_Sql = w_Sql & " ORDER BY NS_STAFFBASISINFO_F.STAFFMNGID"
    '        End If
    '        w_Rs = General.paDBRecordSetOpen(w_Sql)
    '        select_NS_STAFFBASISINFO_F_03 = True
    '    Catch ex As Exception
    '        Throw
    '    End Try
    'End Function
    '2018/02/27 Yamanishi Upd End --------------------------------------------------------------------------------------------------------

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <param name="p_strStaffDataID"></param>
    ''' <returns></returns>
    Public Function select_NS_KINMUIDOINFO_F_01(ByRef w_Rs As ADODB.Recordset,
                                                ByVal p_strStaffDataID As String) As Boolean
        select_NS_KINMUIDOINFO_F_01 = False

        Try

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = ""
                w_Sql = w_Sql & " SELECT "
                w_Sql = w_Sql & "   KINMUDEPTCD "
                w_Sql = w_Sql & " , IDODATE "
                w_Sql = w_Sql & " , ENDDATE "
                w_Sql = w_Sql & " FROM NS_KINMUIDOINFO_F "
                w_Sql = w_Sql & " WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & " AND STAFFMNGID = '" & p_strStaffDataID & "' "
                w_Sql = w_Sql & " ORDER BY IDODATE "

            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = ""
                w_Sql = w_Sql & " SELECT "
                w_Sql = w_Sql & "   KINMUDEPTCD "
                w_Sql = w_Sql & " , IDODATE "
                w_Sql = w_Sql & " , ENDDATE "
                w_Sql = w_Sql & " FROM NS_KINMUIDOINFO_F "
                w_Sql = w_Sql & " WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & " AND STAFFMNGID = '" & p_strStaffDataID & "' "
                w_Sql = w_Sql & " ORDER BY IDODATE "

            End If

            w_Rs = General.paDBRecordSetOpen(w_Sql)
            select_NS_KINMUIDOINFO_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <returns></returns>
    Public Function select_NS_COMMONINFO_M_01(ByRef w_Rs As ADODB.Recordset) As Boolean
        select_NS_COMMONINFO_M_01 = False

        Try

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "SELECT   SYSFROMDAY"
                w_Sql = w_Sql & ",PLANUNIT "
                w_Sql = w_Sql & ",DISPPERIOD "
                w_Sql = w_Sql & ",NAME "
                w_Sql = w_Sql & ",HOPENUM "
                w_Sql = w_Sql & ",HOPENUMDATE "
                w_Sql = w_Sql & ",MULTIOUEN " '2018/02/27 Yamanishi Add
                w_Sql = w_Sql & "FROM NS_COMMONINFO_M "
                w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "SELECT   SYSFROMDAY"
                w_Sql = w_Sql & ",PLANUNIT "
                w_Sql = w_Sql & ",DISPPERIOD "
                w_Sql = w_Sql & ",NAME "
                w_Sql = w_Sql & ",HOPENUM "
                w_Sql = w_Sql & ",HOPENUMDATE "
                w_Sql = w_Sql & ",MULTIOUEN " '2018/02/27 Yamanishi Add
                w_Sql = w_Sql & "FROM NS_COMMONINFO_M "
                w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
            End If

            w_Rs = General.paDBRecordSetOpen(w_Sql)
            select_NS_COMMONINFO_M_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <param name="p_strStaffMngID"></param>
    ''' <returns></returns>
    Public Function select_NS_STAFFMNGHISTORY_F_02(ByRef w_Rs As ADODB.Recordset,
                                                   ByVal p_strStaffMngID As String) As Boolean
        '初期化
        select_NS_STAFFMNGHISTORY_F_02 = False
        Try

            If General.g_InstallType = General.gInstall_Enum.AccessType_PassThrough Then  'ORACLE
                w_Sql = "SELECT EMPDATE, RETIREDATE"
                w_Sql = w_Sql & " FROM NS_STAFFMNGHISTORY_F"
                w_Sql = w_Sql & " WHERE STAFFMNGID = '" & p_strStaffMngID & "'"
                w_Sql = w_Sql & " AND HOSPITALCD = '" & General.g_strHospitalCD & "'"
                w_Sql = w_Sql & " "
            ElseIf General.g_InstallType = General.gInstall_Enum.AccessType_SQL Then 'SQL Server
                w_Sql = "SELECT EMPDATE, RETIREDATE"
                w_Sql = w_Sql & " FROM NS_STAFFMNGHISTORY_F"
                w_Sql = w_Sql & " WHERE STAFFMNGID = '" & p_strStaffMngID & "'"
                w_Sql = w_Sql & " AND HOSPITALCD = '" & General.g_strHospitalCD & "'"
                w_Sql = w_Sql & " "
            End If

            w_Rs = General.paDBRecordSetOpen(w_Sql)
            select_NS_STAFFMNGHISTORY_F_02 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <param name="p_intSortIdx"></param>
    ''' <param name="p_intStartDate"></param>
    ''' <param name="p_intEndDate"></param>
    ''' <returns></returns>
    Public Function select_NS_HOSPBASCHARGES_M_01(ByRef w_Rs As ADODB.Recordset,
                                                  ByVal p_intSortIdx As Integer,
                                                  ByVal p_intStartDate As Integer,
                                                  ByVal p_intEndDate As Integer) As Boolean
        select_NS_HOSPBASCHARGES_M_01 = False

        Try
            w_sqlBuilder = New System.Text.StringBuilder
            w_Sql = String.Empty

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "SELECT AVEINPATIENTNUM , TODOKEDEKBN FROM NS_HOSPBASCHARGES_M"
                w_Sql = w_Sql & " WHERE KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "'"
                If p_intSortIdx = 0 Then
                    w_Sql = w_Sql & " AND FROMDATE <= " & p_intStartDate
                    w_Sql = w_Sql & " AND (TODATE >= " & p_intStartDate
                    w_Sql = w_Sql & " OR TODATE = 0 OR TODATE IS NULL)"
                Else
                    w_Sql = w_Sql & " AND FROMDATE <= " & p_intEndDate
                    w_Sql = w_Sql & " AND (TODATE >= " & p_intEndDate
                    w_Sql = w_Sql & " OR TODATE = 0 OR TODATE IS NULL)"
                End If
                w_Sql = w_Sql & " AND HOSPITALCD = '" & General.g_strHospitalCD & "'"

            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "SELECT AVEINPATIENTNUM , TODOKEDEKBN FROM NS_HOSPBASCHARGES_M"
                w_Sql = w_Sql & " WHERE KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "'"
                If p_intSortIdx = 0 Then
                    w_Sql = w_Sql & " AND FROMDATE <= " & p_intStartDate
                    w_Sql = w_Sql & " AND (TODATE >= " & p_intStartDate
                    w_Sql = w_Sql & " OR TODATE = 0 OR TODATE IS NULL)"
                Else
                    w_Sql = w_Sql & " AND FROMDATE <= " & p_intEndDate
                    w_Sql = w_Sql & " AND (TODATE >= " & p_intEndDate
                    w_Sql = w_Sql & " OR TODATE = 0 OR TODATE IS NULL)"
                End If
                w_Sql = w_Sql & " AND HOSPITALCD = '" & General.g_strHospitalCD & "'"

            End If


            w_Rs = General.paDBRecordSetOpen(w_Sql)

            select_NS_HOSPBASCHARGES_M_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <param name="p_objStaffData"></param>
    ''' <param name="p_intStaffIdx"></param>
    ''' <param name="p_intStDate"></param>
    ''' <param name="p_intEdDate"></param>
    ''' <param name="p_strHolTimeCD"></param>
    ''' <returns></returns>
    Public Function select_NS_NENKYU_F_02(ByRef w_Rs As ADODB.Recordset,
                                          ByVal p_objStaffData As Object,
                                          ByVal p_intStaffIdx As Integer,
                                          ByVal p_intStDate As Integer,
                                          ByVal p_intEdDate As Integer,
                                          ByVal p_strHolTimeCD As String) As Boolean

        select_NS_NENKYU_F_02 = False
        Try

            If General.g_InstallType = General.gInstall_Enum.AccessType_PassThrough Then  'ORACLE
                w_Sql = ""
                w_Sql = w_Sql & "SELECT SUM(NENKYUTIME) "
                w_Sql = w_Sql & "FROM NS_NENKYU_F "
                w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & " AND STAFFMNGID = '" & p_objStaffData(p_intStaffIdx).ID & "' "
                w_Sql = w_Sql & " AND DATEF >= " & p_intStDate & " "
                w_Sql = w_Sql & " AND DATEF <= " & p_intEdDate & " "
                w_Sql = w_Sql & " AND GETCONTENTSKBN = '4' "
                w_Sql = w_Sql & " AND HOLIDAYBUNRUICD = '" & p_strHolTimeCD & "' "
            ElseIf General.g_InstallType = General.gInstall_Enum.AccessType_SQL Then 'SQL Server
                w_Sql = ""
                w_Sql = w_Sql & "SELECT SUM(NENKYUTIME) "
                w_Sql = w_Sql & "FROM NS_NENKYU_F "
                w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & " AND STAFFMNGID = '" & p_objStaffData(p_intStaffIdx).ID & "' "
                w_Sql = w_Sql & " AND DATEF >= " & p_intStDate & " "
                w_Sql = w_Sql & " AND DATEF <= " & p_intEdDate & " "
                w_Sql = w_Sql & " AND GETCONTENTSKBN = '4' "
                w_Sql = w_Sql & " AND HOLIDAYBUNRUICD = '" & p_strHolTimeCD & "' "
            End If

            w_Rs = General.paDBRecordSetOpen(w_Sql)
            select_NS_NENKYU_F_02 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <param name="p_intPlanNo"></param>
    ''' <returns></returns>
    Public Function select_NS_PLANCONTROL_F_01(ByRef w_Rs As ADODB.Recordset,
                                               ByVal p_intPlanNo As Integer) As Boolean
        select_NS_PLANCONTROL_F_01 = False
        Try

            If General.g_InstallType = General.gInstall_Enum.AccessType_PassThrough Then  'ORACLE
                w_Sql = "SELECT PLANPERIODFROM, PLANPERIODTO, PLANDUEDATE, TERM "
                w_Sql = w_Sql & "FROM NS_PLANCONTROL_F "
                w_Sql = w_Sql & "WHERE PLANNO = " & p_intPlanNo & " "
                w_Sql = w_Sql & "AND HOSPITALCD = '" & General.g_strHospitalCD & "' "
            ElseIf General.g_InstallType = General.gInstall_Enum.AccessType_SQL Then 'SQL Server
                w_Sql = "SELECT PLANPERIODFROM, PLANPERIODTO, PLANDUEDATE, TERM "
                w_Sql = w_Sql & "FROM NS_PLANCONTROL_F "
                w_Sql = w_Sql & "WHERE PLANNO = " & p_intPlanNo & " "
                w_Sql = w_Sql & "AND HOSPITALCD = '" & General.g_strHospitalCD & "' "
            End If
            w_Rs = General.paDBRecordSetOpen(w_Sql)
            select_NS_PLANCONTROL_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <param name="p_intKEndlng"></param>
    ''' <param name="p_intYYYYMMDD"></param>
    ''' <returns></returns>
    Public Function select_NS_PLANCONTROL_F_02(ByRef w_Rs As ADODB.Recordset,
                                               ByVal p_intKEndlng As Integer,
                                               ByVal p_intYYYYMMDD As Integer) As Boolean
        select_NS_PLANCONTROL_F_02 = False
        Try
            If General.g_InstallType = General.gInstall_Enum.AccessType_PassThrough Then  'ORACLE
                w_Sql = "SELECT MIN(PLANNO) AS MIN_PLAN FROM NS_PLANCONTROL_F"
                w_Sql = w_Sql & " WHERE HOSPITALCD = '" & General.g_strHospitalCD & "'"
                w_Sql = w_Sql & " AND PLANPERIODFROM <= " & p_intKEndlng
                w_Sql = w_Sql & " AND PLANPERIODTO >= " & p_intYYYYMMDD
            ElseIf General.g_InstallType = General.gInstall_Enum.AccessType_SQL Then 'SQL Server
                w_Sql = "SELECT MIN(PLANNO) AS MIN_PLAN FROM NS_PLANCONTROL_F"
                w_Sql = w_Sql & " WHERE HOSPITALCD = '" & General.g_strHospitalCD & "'"
                w_Sql = w_Sql & " AND PLANPERIODFROM <= " & p_intKEndlng
                w_Sql = w_Sql & " AND PLANPERIODTO >= " & p_intYYYYMMDD
            End If
            w_Rs = General.paDBRecordSetOpen(w_Sql)
            select_NS_PLANCONTROL_F_02 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <param name="p_intPlanNo"></param>
    ''' <returns></returns>
    Public Function select_NS_PLANCONTROL_F_03(ByRef w_Rs As ADODB.Recordset,
                                               ByVal p_intPlanNo As Integer) As Boolean
        select_NS_PLANCONTROL_F_03 = False
        Try
            If General.g_InstallType = General.gInstall_Enum.AccessType_PassThrough Then  'ORACLE
                w_Sql = "SELECT PLANPERIODFROM, PLANPERIODTO, TERM "
                w_Sql = w_Sql & "FROM NS_PLANCONTROL_F "
                w_Sql = w_Sql & "WHERE PLANNO = " & p_intPlanNo & " "
                w_Sql = w_Sql & "AND HOSPITALCD = '" & General.g_strHospitalCD & "' "
            ElseIf General.g_InstallType = General.gInstall_Enum.AccessType_SQL Then 'SQL Server
                w_Sql = "SELECT PLANPERIODFROM, PLANPERIODTO, TERM "
                w_Sql = w_Sql & "FROM NS_PLANCONTROL_F "
                w_Sql = w_Sql & "WHERE PLANNO = " & p_intPlanNo & " "
                w_Sql = w_Sql & "AND HOSPITALCD = '" & General.g_strHospitalCD & "' "
            End If
            w_Rs = General.paDBRecordSetOpen(w_Sql)
            select_NS_PLANCONTROL_F_03 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <param name="p_intPlanNo"></param>
    ''' <returns></returns>
    Public Function select_NS_PLANCONTROL_F_04(ByRef w_Rs As ADODB.Recordset,
                                               ByVal p_intPlanNo As Integer) As Boolean
        select_NS_PLANCONTROL_F_04 = False
        Try
            If General.g_InstallType = General.gInstall_Enum.AccessType_PassThrough Then  'ORACLE
                w_Sql = "SELECT PLANPERIODFROM"
                w_Sql = w_Sql & " FROM NS_PLANCONTROL_F"
                w_Sql = w_Sql & " WHERE PLANNO = " & Convert.ToString(p_intPlanNo)
                w_Sql = w_Sql & " AND HOSPITALCD = '" & General.g_strHospitalCD & "'"
                w_Sql = w_Sql & " "
            ElseIf General.g_InstallType = General.gInstall_Enum.AccessType_SQL Then 'SQL Server
                w_Sql = "SELECT PLANPERIODFROM"
                w_Sql = w_Sql & " FROM NS_PLANCONTROL_F"
                w_Sql = w_Sql & " WHERE PLANNO = " & Convert.ToString(p_intPlanNo)
                w_Sql = w_Sql & " AND HOSPITALCD = '" & General.g_strHospitalCD & "'"
                w_Sql = w_Sql & " "
            End If
            w_Rs = General.paDBRecordSetOpen(w_Sql)
            select_NS_PLANCONTROL_F_04 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <param name="p_strBunruiCD"></param>
    ''' <returns></returns>
    Public Function select_NS_HOLIDAYBUNRUI_M_02(ByRef w_Rs As ADODB.Recordset,
                                                 ByVal p_strBunruiCD As String) As Boolean
        '初期化
        select_NS_HOLIDAYBUNRUI_M_02 = False
        Try

            If General.g_InstallType = General.gInstall_Enum.AccessType_PassThrough Then  'ORACLE
                w_Sql = "Select Name From NS_HOLIDAYBUNRUI_M "
                w_Sql = w_Sql & "Where HolidayBunruiCD = '" & p_strBunruiCD & "' "
                w_Sql = w_Sql & "And HospitalCD = '" & General.g_strHospitalCD & "' "
            ElseIf General.g_InstallType = General.gInstall_Enum.AccessType_SQL Then 'SQL Server
                w_Sql = "Select Name From NS_HOLIDAYBUNRUI_M "
                w_Sql = w_Sql & "Where HolidayBunruiCD = '" & p_strBunruiCD & "' "
                w_Sql = w_Sql & "And HospitalCD = '" & General.g_strHospitalCD & "' "
            End If

            w_Rs = General.paDBRecordSetOpen(w_Sql)
            select_NS_HOLIDAYBUNRUI_M_02 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <param name="p_intPlanNo"></param>
    ''' <returns></returns>
    Public Function select_NS_PLANDECISION_F_01(ByRef w_Rs As ADODB.Recordset,
                                                ByVal p_intPlanNo As Integer) As Boolean
        select_NS_PLANDECISION_F_01 = False

        Try
            If General.g_InstallType = General.gInstall_Enum.AccessType_PassThrough Then  'ORACLE
                w_Sql = "SELECT DECISIONDATE"
                w_Sql = w_Sql & " FROM NS_PLANDECISION_F"
                w_Sql = w_Sql & " WHERE PLANNO = " & Convert.ToString(p_intPlanNo)
                w_Sql = w_Sql & " AND KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "'"
                w_Sql = w_Sql & " AND HOSPITALCD = '" & General.g_strHospitalCD & "'"
                w_Sql = w_Sql & " "
            ElseIf General.g_InstallType = General.gInstall_Enum.AccessType_SQL Then 'SQL Server
                w_Sql = "SELECT DECISIONDATE"
                w_Sql = w_Sql & " FROM NS_PLANDECISION_F"
                w_Sql = w_Sql & " WHERE PLANNO = " & Convert.ToString(p_intPlanNo)
                w_Sql = w_Sql & " AND KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "'"
                w_Sql = w_Sql & " AND HOSPITALCD = '" & General.g_strHospitalCD & "'"
                w_Sql = w_Sql & " "
            End If
            w_Rs = General.paDBRecordSetOpen(w_Sql)
            select_NS_PLANDECISION_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <param name="p_intPlanNo"></param>
    ''' <returns></returns>
    Public Function select_NS_PLANDECISION_F_02(ByRef w_Rs As ADODB.Recordset,
                                                ByVal p_intPlanNo As Integer) As Boolean
        select_NS_PLANDECISION_F_02 = False

        Try

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "SELECT  * "
                w_Sql = w_Sql & "FROM NS_PLANDECISION_F "
                w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "AND KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "' "
                w_Sql = w_Sql & "AND PLANNO = " & p_intPlanNo & " "
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "SELECT  * "
                w_Sql = w_Sql & "FROM NS_PLANDECISION_F "
                w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "AND KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "' "
                w_Sql = w_Sql & "AND PLANNO = " & p_intPlanNo & " "
            End If
            w_Rs = General.paDBRecordSetOpen(w_Sql)
            select_NS_PLANDECISION_F_02 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <param name="p_intDaikyuStartDate"></param>
    ''' <param name="p_intDaikyuEndDate"></param>
    ''' <param name="p_strStaffDataID"></param>
    ''' <returns></returns>
    Public Function select_NS_DAIKYUMNG_F_01(ByRef w_Rs As ADODB.Recordset,
                                             ByVal p_intDaikyuStartDate As Integer,
                                             ByVal p_intDaikyuEndDate As Integer,
                                             ByVal p_strStaffDataID As String,
                                             ByVal p_PlanNo As Integer,
                                             ByVal p_intSelSaveNo As Integer) As Boolean

        select_NS_DAIKYUMNG_F_01 = False

        Try

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                'w_Sql = "SELECT   DM.WORKHOLKINMUDATE "
                'w_Sql = w_Sql & ",DM.WORKHOLKINMUCD "
                'w_Sql = w_Sql & ",DM.GETKBN "
                'w_Sql = w_Sql & ",DM.REGISTFIRSTTIMEDATE "
                'w_Sql = w_Sql & ",DM.LASTUPDTIMEDATE "
                'w_Sql = w_Sql & ",DM.REGISTRANTID "
                'w_Sql = w_Sql & ",DD.SEQ "
                'w_Sql = w_Sql & ",DD.GETFLG "
                'w_Sql = w_Sql & ",DD.GETDAIKYUDATE "
                'w_Sql = w_Sql & ",DD.GETDAIKYUKINMUCD "
                'w_Sql = w_Sql & "FROM NS_DAIKYUMNG_F DM "
                'w_Sql = w_Sql & "LEFT OUTER JOIN NS_DAIKYUDETAILMNG_F DD "
                'w_Sql = w_Sql & "ON    DD.WORKHOLKINMUDATE = DM.WORKHOLKINMUDATE "
                'w_Sql = w_Sql & "AND   DD.STAFFMNGID = DM.STAFFMNGID "
                'w_Sql = w_Sql & "AND   DD.HOSPITALCD = DM.HOSPITALCD "
                'w_Sql = w_Sql & "WHERE DM.WORKHOLKINMUDATE >= " & p_intDaikyuStartDate & " "
                'w_Sql = w_Sql & "AND   DM.WORKHOLKINMUDATE <= " & p_intDaikyuEndDate & " "
                'w_Sql = w_Sql & "AND   DM.STAFFMNGID = '" & p_strStaffDataID & "' "
                'w_Sql = w_Sql & "AND   DM.HOSPITALCD = '" & General.g_strHospitalCD & "' "
                'w_Sql = w_Sql & "ORDER BY DM.WORKHOLKINMUDATE "
                w_Sql = ""
                w_Sql = w_Sql & " SELECT "
                w_Sql = w_Sql & "   DM.WORKHOLKINMUDATE "
                w_Sql = w_Sql & " , DM.WORKHOLKINMUCD "
                w_Sql = w_Sql & " , DM.GETKBN "
                w_Sql = w_Sql & " , DM.REGISTFIRSTTIMEDATE "
                w_Sql = w_Sql & " , DM.LASTUPDTIMEDATE "
                w_Sql = w_Sql & " , DM.REGISTRANTID "
                w_Sql = w_Sql & " , DD.SEQ "
                w_Sql = w_Sql & " , DD.GETFLG "
                w_Sql = w_Sql & " , DD.GETDAIKYUDATE "
                w_Sql = w_Sql & " , DD.GETDAIKYUKINMUCD "
                w_Sql = w_Sql & " FROM NS_DAIKYUMNG_F DM "
                w_Sql = w_Sql & "   LEFT OUTER JOIN NS_DAIKYUDETAILMNG_F DD "
                w_Sql = w_Sql & "     ON  DD.HOSPITALCD = DM.HOSPITALCD "
                w_Sql = w_Sql & "     AND DD.STAFFMNGID = DM.STAFFMNGID "
                w_Sql = w_Sql & "     AND DD.WORKHOLKINMUDATE = DM.WORKHOLKINMUDATE "
                If p_intSelSaveNo > 0 Then
                    w_Sql = w_Sql & "   LEFT OUTER JOIN NS_KINMUIDOINFO_F KI "
                    w_Sql = w_Sql & "     ON  KI.HOSPITALCD  = DM.HOSPITALCD "
                    w_Sql = w_Sql & "     AND KI.STAFFMNGID  = DM.STAFFMNGID "
                    w_Sql = w_Sql & "     AND DM.WORKHOLKINMUDATE BETWEEN KI.IDODATE AND KI.ENDDATE "
                    w_Sql = w_Sql & "   LEFT OUTER JOIN NS_PLANCONTROL_F PC "
                    w_Sql = w_Sql & "     ON  PC.HOSPITALCD  = DM.HOSPITALCD "
                    w_Sql = w_Sql & "     AND DM.WORKHOLKINMUDATE BETWEEN PC.PLANPERIODFROM AND PC.PLANPERIODTO "
                End If
                w_Sql = w_Sql & " WHERE "
                w_Sql = w_Sql & "     DM.HOSPITALCD  = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & " AND DM.STAFFMNGID  = '" & p_strStaffDataID & "' "
                w_Sql = w_Sql & " AND DM.WORKHOLKINMUDATE BETWEEN " & p_intDaikyuStartDate & " AND " & p_intDaikyuEndDate
                If p_intSelSaveNo > 0 Then
                    w_Sql = w_Sql & " AND (KI.KINMUDEPTCD IS NULL "
                    w_Sql = w_Sql & "  OR  KI.KINMUDEPTCD <> '" & General.g_strSelKinmuDeptCD & "' "
                    w_Sql = w_Sql & "  OR  PC.PLANNO      IS NULL "
                    w_Sql = w_Sql & "  OR  PC.PLANNO      <>  " & p_PlanNo & ") "
                    w_Sql = w_Sql & " UNION ALL "
                    w_Sql = w_Sql & " SELECT "
                    w_Sql = w_Sql & "   DM.WORKHOLKINMUDATE "
                    w_Sql = w_Sql & " , DM.WORKHOLKINMUCD "
                    w_Sql = w_Sql & " , DM.GETKBN "
                    w_Sql = w_Sql & " , DM.REGISTFIRSTTIMEDATE "
                    w_Sql = w_Sql & " , DM.LASTUPDTIMEDATE "
                    w_Sql = w_Sql & " , DM.REGISTRANTID "
                    w_Sql = w_Sql & " , NULL AS SEQ "
                    w_Sql = w_Sql & " , NULL AS GETFLG "
                    w_Sql = w_Sql & " , NULL AS GETDAIKYUDATE "
                    w_Sql = w_Sql & " , NULL AS GETDAIKYUKINMUCD "
                    w_Sql = w_Sql & " FROM NS_TEMPDAIKYUMNG_F DM "
                    w_Sql = w_Sql & "   INNER JOIN NS_KINMUIDOINFO_F KI "
                    w_Sql = w_Sql & "     ON  KI.HOSPITALCD  = DM.HOSPITALCD "
                    w_Sql = w_Sql & "     AND KI.STAFFMNGID  = DM.STAFFMNGID "
                    w_Sql = w_Sql & "     AND DM.WORKHOLKINMUDATE BETWEEN KI.IDODATE AND KI.ENDDATE "
                    w_Sql = w_Sql & "     AND KI.KINMUDEPTCD = DM.KINMUDEPTCD "
                    w_Sql = w_Sql & "   INNER JOIN NS_PLANCONTROL_F PC "
                    w_Sql = w_Sql & "     ON  PC.HOSPITALCD  = DM.HOSPITALCD "
                    w_Sql = w_Sql & "     AND PC.PLANNO      = DM.PLANNO "
                    w_Sql = w_Sql & "     AND DM.WORKHOLKINMUDATE BETWEEN PC.PLANPERIODFROM AND PC.PLANPERIODTO "
                    w_Sql = w_Sql & " WHERE "
                    w_Sql = w_Sql & "     DM.HOSPITALCD  = '" & General.g_strHospitalCD & "' "
                    w_Sql = w_Sql & " AND DM.PLANNO      =  " & p_PlanNo
                    w_Sql = w_Sql & " AND DM.KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "' "
                    w_Sql = w_Sql & " AND DM.SAVENO      =  " & p_intSelSaveNo
                    w_Sql = w_Sql & " AND DM.STAFFMNGID  = '" & p_strStaffDataID & "' "
                    w_Sql = w_Sql & " AND DM.WORKHOLKINMUDATE BETWEEN " & p_intDaikyuStartDate & " AND " & p_intDaikyuEndDate
                End If
                w_Sql = w_Sql & " ORDER BY "
                w_Sql = w_Sql & "   WORKHOLKINMUDATE "
                w_Sql = w_Sql & " , GETDAIKYUDATE "
                w_Sql = w_Sql & " , SEQ "
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                'w_Sql = "SELECT   DM.WORKHOLKINMUDATE "
                'w_Sql = w_Sql & ",DM.WORKHOLKINMUCD "
                'w_Sql = w_Sql & ",DM.GETKBN "
                'w_Sql = w_Sql & ",DM.REGISTFIRSTTIMEDATE "
                'w_Sql = w_Sql & ",DM.LASTUPDTIMEDATE "
                'w_Sql = w_Sql & ",DM.REGISTRANTID "
                'w_Sql = w_Sql & ",DD.SEQ "
                'w_Sql = w_Sql & ",DD.GETFLG "
                'w_Sql = w_Sql & ",DD.GETDAIKYUDATE "
                'w_Sql = w_Sql & ",DD.GETDAIKYUKINMUCD "
                'w_Sql = w_Sql & "FROM NS_DAIKYUMNG_F DM "
                'w_Sql = w_Sql & "LEFT OUTER JOIN NS_DAIKYUDETAILMNG_F DD "
                'w_Sql = w_Sql & "ON    DD.WORKHOLKINMUDATE = DM.WORKHOLKINMUDATE "
                'w_Sql = w_Sql & "AND   DD.STAFFMNGID = DM.STAFFMNGID "
                'w_Sql = w_Sql & "AND   DD.HOSPITALCD = DM.HOSPITALCD "
                'w_Sql = w_Sql & "WHERE DM.WORKHOLKINMUDATE >= " & p_intDaikyuStartDate & " "
                'w_Sql = w_Sql & "AND   DM.WORKHOLKINMUDATE <= " & p_intDaikyuEndDate & " "
                'w_Sql = w_Sql & "AND   DM.STAFFMNGID = '" & p_strStaffDataID & "' "
                'w_Sql = w_Sql & "AND   DM.HOSPITALCD = '" & General.g_strHospitalCD & "' "
                'w_Sql = w_Sql & "ORDER BY DM.WORKHOLKINMUDATE "
                w_Sql = ""
                w_Sql = w_Sql & " SELECT "
                w_Sql = w_Sql & "   DM.WORKHOLKINMUDATE "
                w_Sql = w_Sql & " , DM.WORKHOLKINMUCD "
                w_Sql = w_Sql & " , DM.GETKBN "
                w_Sql = w_Sql & " , DM.REGISTFIRSTTIMEDATE "
                w_Sql = w_Sql & " , DM.LASTUPDTIMEDATE "
                w_Sql = w_Sql & " , DM.REGISTRANTID "
                w_Sql = w_Sql & " , DD.SEQ "
                w_Sql = w_Sql & " , DD.GETFLG "
                w_Sql = w_Sql & " , DD.GETDAIKYUDATE "
                w_Sql = w_Sql & " , DD.GETDAIKYUKINMUCD "
                w_Sql = w_Sql & " FROM NS_DAIKYUMNG_F DM "
                w_Sql = w_Sql & "   LEFT OUTER JOIN NS_DAIKYUDETAILMNG_F DD "
                w_Sql = w_Sql & "     ON  DD.HOSPITALCD = DM.HOSPITALCD "
                w_Sql = w_Sql & "     AND DD.STAFFMNGID = DM.STAFFMNGID "
                w_Sql = w_Sql & "     AND DD.WORKHOLKINMUDATE = DM.WORKHOLKINMUDATE "
                If p_intSelSaveNo > 0 Then
                    w_Sql = w_Sql & "   LEFT OUTER JOIN NS_KINMUIDOINFO_F KI "
                    w_Sql = w_Sql & "     ON  KI.HOSPITALCD  = DM.HOSPITALCD "
                    w_Sql = w_Sql & "     AND KI.STAFFMNGID  = DM.STAFFMNGID "
                    w_Sql = w_Sql & "     AND DM.WORKHOLKINMUDATE BETWEEN KI.IDODATE AND KI.ENDDATE "
                    w_Sql = w_Sql & "   LEFT OUTER JOIN NS_PLANCONTROL_F PC "
                    w_Sql = w_Sql & "     ON  PC.HOSPITALCD  = DM.HOSPITALCD "
                    w_Sql = w_Sql & "     AND DM.WORKHOLKINMUDATE BETWEEN PC.PLANPERIODFROM AND PC.PLANPERIODTO "
                End If
                w_Sql = w_Sql & " WHERE "
                w_Sql = w_Sql & "     DM.HOSPITALCD  = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & " AND DM.STAFFMNGID  = '" & p_strStaffDataID & "' "
                w_Sql = w_Sql & " AND DM.WORKHOLKINMUDATE BETWEEN " & p_intDaikyuStartDate & " AND " & p_intDaikyuEndDate
                If p_intSelSaveNo > 0 Then
                    w_Sql = w_Sql & " AND (KI.KINMUDEPTCD IS NULL "
                    w_Sql = w_Sql & "  OR  KI.KINMUDEPTCD <> '" & General.g_strSelKinmuDeptCD & "' "
                    w_Sql = w_Sql & "  OR  PC.PLANNO      IS NULL "
                    w_Sql = w_Sql & "  OR  PC.PLANNO      <>  " & p_PlanNo & ") "
                    w_Sql = w_Sql & " UNION ALL "
                    w_Sql = w_Sql & " SELECT "
                    w_Sql = w_Sql & "   DM.WORKHOLKINMUDATE "
                    w_Sql = w_Sql & " , DM.WORKHOLKINMUCD "
                    w_Sql = w_Sql & " , DM.GETKBN "
                    w_Sql = w_Sql & " , DM.REGISTFIRSTTIMEDATE "
                    w_Sql = w_Sql & " , DM.LASTUPDTIMEDATE "
                    w_Sql = w_Sql & " , DM.REGISTRANTID "
                    w_Sql = w_Sql & " , NULL AS SEQ "
                    w_Sql = w_Sql & " , NULL AS GETFLG "
                    w_Sql = w_Sql & " , NULL AS GETDAIKYUDATE "
                    w_Sql = w_Sql & " , NULL AS GETDAIKYUKINMUCD "
                    w_Sql = w_Sql & " FROM NS_TEMPDAIKYUMNG_F DM "
                    w_Sql = w_Sql & "   INNER JOIN NS_KINMUIDOINFO_F KI "
                    w_Sql = w_Sql & "     ON  KI.HOSPITALCD  = DM.HOSPITALCD "
                    w_Sql = w_Sql & "     AND KI.STAFFMNGID  = DM.STAFFMNGID "
                    w_Sql = w_Sql & "     AND DM.WORKHOLKINMUDATE BETWEEN KI.IDODATE AND KI.ENDDATE "
                    w_Sql = w_Sql & "     AND KI.KINMUDEPTCD = DM.KINMUDEPTCD "
                    w_Sql = w_Sql & "   INNER JOIN NS_PLANCONTROL_F PC "
                    w_Sql = w_Sql & "     ON  PC.HOSPITALCD  = DM.HOSPITALCD "
                    w_Sql = w_Sql & "     AND PC.PLANNO      = DM.PLANNO "
                    w_Sql = w_Sql & "     AND DM.WORKHOLKINMUDATE BETWEEN PC.PLANPERIODFROM AND PC.PLANPERIODTO "
                    w_Sql = w_Sql & " WHERE "
                    w_Sql = w_Sql & "     DM.HOSPITALCD  = '" & General.g_strHospitalCD & "' "
                    w_Sql = w_Sql & " AND DM.PLANNO      =  " & p_PlanNo
                    w_Sql = w_Sql & " AND DM.KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "' "
                    w_Sql = w_Sql & " AND DM.SAVENO      =  " & p_intSelSaveNo
                    w_Sql = w_Sql & " AND DM.STAFFMNGID  = '" & p_strStaffDataID & "' "
                    w_Sql = w_Sql & " AND DM.WORKHOLKINMUDATE BETWEEN " & p_intDaikyuStartDate & " AND " & p_intDaikyuEndDate
                End If
                w_Sql = w_Sql & " ORDER BY "
                w_Sql = w_Sql & "   WORKHOLKINMUDATE "
                w_Sql = w_Sql & " , GETDAIKYUDATE "
                w_Sql = w_Sql & " , SEQ "
            End If

            w_Rs = General.paDBRecordSetOpen(w_Sql)
            select_NS_DAIKYUMNG_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HJ
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <param name="p_intSelDate"></param>
    ''' <param name="p_strSelDate"></param>
    ''' <param name="p_strMngStaffID"></param>
    ''' <param name="p_strHospitalCD"></param>
    ''' <returns></returns>
    Public Function select_NS_DAIKYUDETAILMNG_F_01(ByRef w_Rs As ADODB.Recordset,
                                                   ByVal p_intSelDate As Integer,
                                                   ByVal p_strSelDate As String,
                                                   ByVal p_strMngStaffID As String,
                                                   ByVal p_strHospitalCD As String) As Boolean
        Try
            select_NS_DAIKYUDETAILMNG_F_01 = False
            w_sqlBuilder = New System.Text.StringBuilder

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then 'ORACLE
                w_sqlBuilder.Append("select * from NS_DAIKYUDETAILMNG_F")
                w_sqlBuilder.Append(" where  WorkHolKinmuDate = ")
                w_sqlBuilder.Append(p_intSelDate)
                w_sqlBuilder.Append(" and GETDAIKYUDATE <> ")
                w_sqlBuilder.Append(p_strSelDate)
                w_sqlBuilder.Append(" and StaffMngID = '")
                w_sqlBuilder.Append(p_strMngStaffID)
                w_sqlBuilder.Append("' and HospitalCD = '")
                w_sqlBuilder.Append(p_strHospitalCD)
                w_sqlBuilder.Append("'")
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then 'SQL Server
                w_sqlBuilder.Append("select * from NS_DAIKYUDETAILMNG_F")
                w_sqlBuilder.Append(" where  WorkHolKinmuDate = ")
                w_sqlBuilder.Append(p_intSelDate)
                w_sqlBuilder.Append(" and GETDAIKYUDATE <> ")
                w_sqlBuilder.Append(p_strSelDate)
                w_sqlBuilder.Append(" and StaffMngID = '")
                w_sqlBuilder.Append(p_strMngStaffID)
                w_sqlBuilder.Append("' and HospitalCD = '")
                w_sqlBuilder.Append(p_strHospitalCD)
                w_sqlBuilder.Append("'")
            End If
            w_Sql = w_sqlBuilder.ToString
            w_Rs = General.paDBRecordSetOpen(w_Sql)
            select_NS_DAIKYUDETAILMNG_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <param name="p_strStaffDataID"></param>
    ''' <param name="p_intRuikeiStartNo"></param>
    ''' <param name="p_intRuikeiEndNo"></param>
    ''' <returns></returns>
    Public Function select_NS_RUIKEITIME_F_01(ByRef w_Rs As ADODB.Recordset,
                                              ByVal p_strStaffDataID As String,
                                              ByVal p_intRuikeiStartNo As Integer,
                                              ByVal p_intRuikeiEndNo As Integer) As Boolean
        select_NS_RUIKEITIME_F_01 = False

        Try
            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "SELECT KINMUTIME FROM NS_RUIKEITIME_F"
                w_Sql = w_Sql & " WHERE STAFFMNGID = '" & p_strStaffDataID & "'"
                w_Sql = w_Sql & " AND PLANNO >= " & p_intRuikeiStartNo
                w_Sql = w_Sql & " AND PLANNO <= " & p_intRuikeiEndNo
                w_Sql = w_Sql & " AND KINMUKBN = '2'"
                w_Sql = w_Sql & " AND HOSPITALCD = '" & General.g_strHospitalCD & "'"
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "SELECT KINMUTIME FROM NS_RUIKEITIME_F"
                w_Sql = w_Sql & " WHERE STAFFMNGID = '" & p_strStaffDataID & "'"
                w_Sql = w_Sql & " AND PLANNO >= " & p_intRuikeiStartNo
                w_Sql = w_Sql & " AND PLANNO <= " & p_intRuikeiEndNo
                w_Sql = w_Sql & " AND KINMUKBN = '2'"
                w_Sql = w_Sql & " AND HOSPITALCD = '" & General.g_strHospitalCD & "'"
            End If

            w_Rs = General.paDBRecordSetOpen(w_Sql)
            select_NS_RUIKEITIME_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HC
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <returns></returns>
    Public Function select_NS_PACKAGE_M_01(ByRef w_Rs As ADODB.Recordset) As Boolean

        '初期化
        select_NS_PACKAGE_M_01 = False

        Try
            w_sqlBuilder = New System.Text.StringBuilder

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = ""
                w_Sql = "Select PACKAGECD, USEFLG"
                w_Sql = w_Sql & " From NS_PACKAGE_M"
                w_Sql = w_Sql & " Where HospitalCD = '" & General.g_strHospitalCD & "'"
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = ""
                w_Sql = "Select PACKAGECD, USEFLG"
                w_Sql = w_Sql & " From NS_PACKAGE_M"
                w_Sql = w_Sql & " Where HospitalCD = '" & General.g_strHospitalCD & "'"
            End If

            w_Rs = General.paDBRecordSetOpen(w_Sql)
            select_NS_PACKAGE_M_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <param name="w_CurrentDate"></param>
    ''' <param name="w_StaffID"></param>
    ''' <returns></returns>
    Public Function select_NS_KINMUPLAN_F_01(ByRef w_Rs As ADODB.Recordset,
                                             ByVal w_CurrentDate As Integer,
                                             ByVal w_StaffID As String) As Boolean
        select_NS_KINMUPLAN_F_01 = False

        Try

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "Select LastUpdTimeDate From NS_KINMUPLAN_F "
                w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "AND DATEF        =  " & w_CurrentDate & " "
                w_Sql = w_Sql & "AND STAFFMNGID   = '" & w_StaffID & "' "
                w_Rs = General.paDBRecordSetOpen(w_Sql)
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "Select LastUpdTimeDate From NS_KINMUPLAN_F "
                w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "AND DATEF        =  " & w_CurrentDate & " "
                w_Sql = w_Sql & "AND STAFFMNGID   = '" & w_StaffID & "' "
                w_Rs = General.paDBRecordSetOpen(w_Sql)
            End If

            w_Rs = General.paDBRecordSetOpen(w_Sql)
            select_NS_KINMUPLAN_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <param name="p_intCurrentDate"></param>
    ''' <param name="p_strStaffID"></param>
    ''' <returns></returns>
    Public Function select_NS_KINMUCHGHISTORY_F_01(ByRef w_Rs As ADODB.Recordset,
                                                   ByVal p_intCurrentDate As Integer,
                                                   ByVal p_strStaffID As String) As Boolean
        select_NS_KINMUCHGHISTORY_F_01 = False

        Try
            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "Select  SEQ "
                w_Sql = w_Sql & "From NS_KINMUCHGHISTORY_F "
                w_Sql = w_Sql & "Where HospitalCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "And DateF = " & p_intCurrentDate & " "
                w_Sql = w_Sql & "And StaffMngID = '" & p_strStaffID & "' "
                w_Sql = w_Sql & "Order By SEQ Desc "
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "Select  SEQ "
                w_Sql = w_Sql & "From NS_KINMUCHGHISTORY_F "
                w_Sql = w_Sql & "Where HospitalCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "And DateF = " & p_intCurrentDate & " "
                w_Sql = w_Sql & "And StaffMngID = '" & p_strStaffID & "' "
                w_Sql = w_Sql & "Order By SEQ Desc "
            End If

            w_Rs = General.paDBRecordSetOpen(w_Sql)
            select_NS_KINMUCHGHISTORY_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <returns></returns>
    Public Function select_NS_KINMUNAME_M_01(ByRef w_Rs As ADODB.Recordset) As Boolean
        select_NS_KINMUNAME_M_01 = False

        Try
            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "Select MarkF, KinmuCD, AllocBunruiCD "
                w_Sql = w_Sql & "From NS_KINMUNAME_M "
                w_Sql = w_Sql & "Where AllocFlg = '2' "
                w_Sql = w_Sql & "And HospitalCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "Order By DispNo "
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "Select MarkF, KinmuCD, AllocBunruiCD "
                w_Sql = w_Sql & "From NS_KINMUNAME_M "
                w_Sql = w_Sql & "Where AllocFlg = '2' "
                w_Sql = w_Sql & "And HospitalCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "Order By DispNo "
            End If
            w_Rs = General.paDBRecordSetOpen(w_Sql)
            select_NS_KINMUNAME_M_01 = False
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HD
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <returns></returns>
    Public Function select_NS_KINMUNAME_M_02(ByRef w_Rs As ADODB.Recordset) As Boolean
        Try
            select_NS_KINMUNAME_M_02 = False
            w_sqlBuilder = New System.Text.StringBuilder

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then 'ORACLE
                w_sqlBuilder.Append("SELECT   KINMUCD ")
                w_sqlBuilder.Append(",MARKF ")
                w_sqlBuilder.Append(",NAME ")
                w_sqlBuilder.Append(",HOLIDAYBUNRUICD ")
                w_sqlBuilder.Append(",EFFTODATE ")
                w_sqlBuilder.Append("FROM NS_KINMUNAME_M ")
                w_sqlBuilder.Append("WHERE HOSPITALCD = '")
                w_sqlBuilder.Append(General.g_strHospitalCD)
                w_sqlBuilder.Append("' ORDER BY DISPNO ")
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then 'SQL Server
                w_sqlBuilder.Append("SELECT   KINMUCD ")
                w_sqlBuilder.Append(",MARKF ")
                w_sqlBuilder.Append(",NAME ")
                w_sqlBuilder.Append(",HOLIDAYBUNRUICD ")
                w_sqlBuilder.Append(",EFFTODATE ")
                w_sqlBuilder.Append("FROM NS_KINMUNAME_M ")
                w_sqlBuilder.Append("WHERE HOSPITALCD = '")
                w_sqlBuilder.Append(General.g_strHospitalCD)
                w_sqlBuilder.Append("' ORDER BY DISPNO ")
            End If

            w_Sql = w_sqlBuilder.ToString
            w_Rs = General.paDBRecordSetOpen(w_Sql)

            select_NS_KINMUNAME_M_02 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HH
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <returns></returns>
    Public Function select_NS_KINMUNAME_M_03(ByRef w_Rs As ADODB.Recordset) As Boolean
        Try
            select_NS_KINMUNAME_M_03 = False
            w_sqlBuilder = New System.Text.StringBuilder

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then 'ORACLE
                w_sqlBuilder.Append("Select KinmuCD, Name, AllocBunruiCD ")
                w_sqlBuilder.Append("From NS_KINMUNAME_M ")
                w_sqlBuilder.Append("Where AllocFlg = '2' ")
                w_sqlBuilder.Append("And HospitalCD = '")
                w_sqlBuilder.Append(General.g_strHospitalCD)
                w_sqlBuilder.Append("' Order By DispNo ")
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then 'SQL Server
                w_sqlBuilder.Append("Select KinmuCD, Name, AllocBunruiCD ")
                w_sqlBuilder.Append("From NS_KINMUNAME_M ")
                w_sqlBuilder.Append("Where AllocFlg = '2' ")
                w_sqlBuilder.Append("And HospitalCD = '")
                w_sqlBuilder.Append(General.g_strHospitalCD)
                w_sqlBuilder.Append("' Order By DispNo ")
            End If

            w_Sql = w_sqlBuilder.ToString
            w_Rs = General.paDBRecordSetOpen(w_Sql)

            select_NS_KINMUNAME_M_03 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HJ
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <returns></returns>
    Public Function select_NS_KINMUNAME_M_04(ByRef w_Rs As ADODB.Recordset) As Boolean
        Try
            select_NS_KINMUNAME_M_04 = False
            w_sqlBuilder = New System.Text.StringBuilder

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then 'ORACLE
                w_sqlBuilder.Append("select KinmuCD, Name, MarkF from NS_KINMUNAME_M")
                w_sqlBuilder.Append(" where GetDaikyuFlg = '1'")
                w_sqlBuilder.Append(" and HospitalCD = '")
                w_sqlBuilder.Append(General.g_strHospitalCD)
                w_sqlBuilder.Append("'")
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then 'SQL Server
                w_sqlBuilder.Append("select KinmuCD, Name, MarkF from NS_KINMUNAME_M")
                w_sqlBuilder.Append(" where GetDaikyuFlg = '1'")
                w_sqlBuilder.Append(" and HospitalCD = '")
                w_sqlBuilder.Append(General.g_strHospitalCD)
                w_sqlBuilder.Append("'")
            End If
            w_Sql = w_sqlBuilder.ToString
            w_Rs = General.paDBRecordSetOpen(w_Sql)
            select_NS_KINMUNAME_M_04 = True
        Catch ex As Exception
            Throw
        End Try
    End Function
    '2018/02/28 Yamanshi Del Start ----------------------------------------------------------------------------------------------
    '''' <summary>
    '''' NSK0000HA
    '''' </summary>
    '''' <param name="w_Rs"></param>
    '''' <param name="p_intKinmuTimeMCol_Min"></param>
    '''' <param name="p_intKinmuTimeMCol_Max"></param>
    '''' <returns></returns>
    'Public Function select_NS_KINMUTIME_M_01(ByRef w_Rs As ADODB.Recordset,
    '                                         ByVal p_intKinmuTimeMCol_Min As Integer,
    '                                         ByVal p_intKinmuTimeMCol_Max As Integer) As Boolean
    '    select_NS_KINMUTIME_M_01 = False

    '    Try

    '        If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
    '            w_Sql = ""
    '            w_Sql = w_Sql & " SELECT "
    '            w_Sql = w_Sql & "   EMPCD "       '採用CD
    '            w_Sql = w_Sql & " , KINMUDEPTCD " '勤務部署CD
    '            w_Sql = w_Sql & " , KINMUCD "     '勤務CD
    '            For i As gKinmuTimeMCol = p_intKinmuTimeMCol_Min To p_intKinmuTimeMCol_Max
    '                w_Sql = w_Sql & " , " & i.ToString("G")
    '            Next i
    '            w_Sql = w_Sql & " FROM NS_KINMUTIME_M "
    '            w_Sql = w_Sql & " WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
    '        ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
    '            w_Sql = ""
    '            w_Sql = w_Sql & " SELECT "
    '            w_Sql = w_Sql & "   EMPCD "       '採用CD
    '            w_Sql = w_Sql & " , KINMUDEPTCD " '勤務部署CD
    '            w_Sql = w_Sql & " , KINMUCD "     '勤務CD
    '            For i As gKinmuTimeMCol = p_intKinmuTimeMCol_Min To p_intKinmuTimeMCol_Max
    '                w_Sql = w_Sql & " , " & i.ToString("G")
    '            Next i
    '            w_Sql = w_Sql & " FROM NS_KINMUTIME_M "
    '            w_Sql = w_Sql & " WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
    '        End If
    '        w_Rs = General.paDBRecordSetOpen(w_Sql)
    '        select_NS_KINMUTIME_M_01 = True
    '    Catch ex As Exception
    '        Throw
    '    End Try
    'End Function
    '2018/02/28 Yamanshi Del End ------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <returns></returns>
    Public Function select_NS_KEYBOARD_F_01(ByRef w_Rs As ADODB.Recordset) As Boolean
        select_NS_KEYBOARD_F_01 = False

        Try

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = ""
                w_Sql = w_Sql & "SELECT "
                w_Sql = w_Sql & "    ALLOCKBN, "
                w_Sql = w_Sql & "    KINMUCD1, "
                w_Sql = w_Sql & "    KINMUCD2 "
                w_Sql = w_Sql & "FROM "
                w_Sql = w_Sql & "    NS_KINMUPATTERN_F "
                w_Sql = w_Sql & "WHERE "
                w_Sql = w_Sql & "    HOSPITALCD = '" & General.g_strHospitalCD & "' AND "
                w_Sql = w_Sql & "    KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "' AND "
                w_Sql = w_Sql & "    ALLOCKBN IN ('3', '4', '5') "
                w_Sql = w_Sql & "ORDER BY "
                w_Sql = w_Sql & "    SEQ "
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = ""
                w_Sql = w_Sql & "SELECT "
                w_Sql = w_Sql & "    ALLOCKBN, "
                w_Sql = w_Sql & "    KINMUCD1, "
                w_Sql = w_Sql & "    KINMUCD2 "
                w_Sql = w_Sql & "FROM "
                w_Sql = w_Sql & "    NS_KINMUPATTERN_F "
                w_Sql = w_Sql & "WHERE "
                w_Sql = w_Sql & "    HOSPITALCD = '" & General.g_strHospitalCD & "' AND "
                w_Sql = w_Sql & "    KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "' AND "
                w_Sql = w_Sql & "    ALLOCKBN IN ('3', '4', '5') "
                w_Sql = w_Sql & "ORDER BY "
                w_Sql = w_Sql & "    SEQ "
            End If
            w_Rs = General.paDBRecordSetOpen(w_Sql)
            select_NS_KEYBOARD_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <param name="p_intSDate"></param>
    ''' <param name="p_intEDate"></param>
    ''' <param name="p_objStaffData"></param>
    ''' <param name="p_intStaffIdx"></param>
    ''' <param name="p_strOneDayKinmuCD"></param>
    ''' <param name="p_strHalfDayKinmuCD"></param>
    ''' <returns></returns>
    Public Function select_NS_KINMURESULT_F_01(ByRef w_Rs As ADODB.Recordset,
                                               ByVal p_intSDate As Integer,
                                               ByVal p_intEDate As Integer,
                                               ByVal p_objStaffData As Object,
                                               ByVal p_intStaffIdx As Integer,
                                               ByVal p_strOneDayKinmuCD As String,
                                               ByVal p_strHalfDayKinmuCD As String) As Boolean
        select_NS_KINMURESULT_F_01 = False

        Try

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = ""
                w_Sql = w_Sql & "SELECT ( "
                w_Sql = w_Sql & "	SELECT COUNT(*) "
                w_Sql = w_Sql & "	FROM NS_KINMURESULT_F "
                w_Sql = w_Sql & "	WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "		AND DATEF >= " & p_intSDate & " "
                w_Sql = w_Sql & "		AND DATEF <= " & p_intEDate & " "
                w_Sql = w_Sql & "		AND STAFFMNGID = '" & p_objStaffData(p_intStaffIdx).ID & "' "
                w_Sql = w_Sql & "		AND KINMUCD IN (" & p_strOneDayKinmuCD & ") "
                w_Sql = w_Sql & ") + ( "
                w_Sql = w_Sql & "	SELECT COUNT(*) "
                w_Sql = w_Sql & "	FROM NS_KINMURESULT_F "
                w_Sql = w_Sql & "	WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "		AND DATEF >= " & p_intSDate & " "
                w_Sql = w_Sql & "		AND DATEF <= " & p_intEDate & " "
                w_Sql = w_Sql & "		AND STAFFMNGID = '" & p_objStaffData(p_intStaffIdx).ID & "' "
                w_Sql = w_Sql & "		AND KINMUCD IN (" & p_strHalfDayKinmuCD & ") "
                w_Sql = w_Sql & ") * 0.5 "
                w_Sql = w_Sql & "FROM DUAL "
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = ""
                w_Sql = w_Sql & "SELECT ( "
                w_Sql = w_Sql & "	SELECT COUNT(*) "
                w_Sql = w_Sql & "	FROM NS_KINMURESULT_F "
                w_Sql = w_Sql & "	WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "		AND DATEF >= " & p_intSDate & " "
                w_Sql = w_Sql & "		AND DATEF <= " & p_intEDate & " "
                w_Sql = w_Sql & "		AND STAFFMNGID = '" & p_objStaffData(p_intStaffIdx).ID & "' "
                w_Sql = w_Sql & "		AND KINMUCD IN (" & p_strOneDayKinmuCD & ") "
                w_Sql = w_Sql & ") + ( "
                w_Sql = w_Sql & "	SELECT COUNT(*) "
                w_Sql = w_Sql & "	FROM NS_KINMURESULT_F "
                w_Sql = w_Sql & "	WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "		AND DATEF >= " & p_intSDate & " "
                w_Sql = w_Sql & "		AND DATEF <= " & p_intEDate & " "
                w_Sql = w_Sql & "		AND STAFFMNGID = '" & p_objStaffData(p_intStaffIdx).ID & "' "
                w_Sql = w_Sql & "		AND KINMUCD IN (" & p_strHalfDayKinmuCD & ") "
                w_Sql = w_Sql & ") * 0.5 "
            End If
            w_Rs = General.paDBRecordSetOpen(w_Sql)
            select_NS_KINMURESULT_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HB, NSK0000HC
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <returns></returns>
    Public Function select_NS_SETKINMU_M_01(ByRef w_Rs As ADODB.Recordset) As Boolean

        Try
            select_NS_SETKINMU_M_01 = False
            w_sqlBuilder = New System.Text.StringBuilder

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then 'ORACLE
                w_sqlBuilder.Append("SELECT * FROM NS_SETKINMU_M ")
                w_sqlBuilder.Append("WHERE HOSPITALCD = '")
                w_sqlBuilder.Append(General.g_strHospitalCD)
                w_sqlBuilder.Append("' AND KINMUDEPTCD = '")
                w_sqlBuilder.Append(General.g_strSelKinmuDeptCD)
                w_sqlBuilder.Append("' ORDER BY DISPNO ")
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then 'SQL Server
                w_sqlBuilder.Append("SELECT * FROM NS_SETKINMU_M ")
                w_sqlBuilder.Append("WHERE HOSPITALCD = '")
                w_sqlBuilder.Append(General.g_strHospitalCD)
                w_sqlBuilder.Append("' AND KINMUDEPTCD = '")
                w_sqlBuilder.Append(General.g_strSelKinmuDeptCD)
                w_sqlBuilder.Append("' ORDER BY DISPNO ")
            End If

            w_Sql = w_sqlBuilder.ToString
            w_Rs = General.paDBRecordSetOpen(w_Sql)
            select_NS_SETKINMU_M_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HD
    ''' </summary>
    ''' <param name="w_Rs"></param>
    ''' <returns></returns>
    Public Function select_NS_SETKINMUNAME_F_01(ByRef w_Rs As ADODB.Recordset) As Boolean
        Try
            select_NS_SETKINMUNAME_F_01 = False
            w_sqlBuilder = New System.Text.StringBuilder

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then 'ORACLE
                w_sqlBuilder.Append("SELECT   KINMUCD ")
                w_sqlBuilder.Append("FROM NS_SETKINMUNAME_F ")
                w_sqlBuilder.Append("WHERE HOSPITALCD = '")
                w_sqlBuilder.Append(General.g_strHospitalCD)
                w_sqlBuilder.Append("' AND KINMUDEPTCD = '")
                w_sqlBuilder.Append(General.g_strSelKinmuDeptCD)
                w_sqlBuilder.Append("' ORDER BY DISPNO ")
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then 'SQL Server
                w_sqlBuilder.Append("SELECT   KINMUCD ")
                w_sqlBuilder.Append("FROM NS_SETKINMUNAME_F ")
                w_sqlBuilder.Append("WHERE HOSPITALCD = '")
                w_sqlBuilder.Append(General.g_strHospitalCD)
                w_sqlBuilder.Append("' AND KINMUDEPTCD = '")
                w_sqlBuilder.Append(General.g_strSelKinmuDeptCD)
                w_sqlBuilder.Append("' ORDER BY DISPNO ")
            End If

            w_Sql = w_sqlBuilder.ToString
            w_Rs = General.paDBRecordSetOpen(w_Sql)

            select_NS_SETKINMUNAME_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="p_intCurrentDate"></param>
    ''' <param name="p_strStaffDataID"></param>
    ''' <param name="p_strKinmuCD"></param>
    ''' <param name="p_strRiyuKBN"></param>
    ''' <param name="p_strComment"></param>
    ''' <param name="p_strSysDate"></param>
    ''' <returns></returns>
    Public Function insert_NS_KINMUPLAN_F_01(ByVal p_intCurrentDate As Integer,
                                             ByVal p_strStaffDataID As String,
                                             ByVal p_strKinmuCD As String,
                                             ByVal p_strRiyuKBN As String,
                                             ByVal p_strComment As String,
                                             ByVal p_strSysDate As String) As Boolean
        insert_NS_KINMUPLAN_F_01 = False

        Try
            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "Insert Into NS_KINMUPLAN_F ("
                w_Sql = w_Sql & " HospitalCD"
                w_Sql = w_Sql & ", DateF"
                w_Sql = w_Sql & ", StaffMngID"
                w_Sql = w_Sql & ", KinmuCD"
                w_Sql = w_Sql & ", ReasonKbn"
                w_Sql = w_Sql & ", HOPECOMMENT" '2015/04/10 Bando Add
                w_Sql = w_Sql & ", RegistFirstTimeDate"
                w_Sql = w_Sql & ", LastUpdTimeDate"
                w_Sql = w_Sql & ", RegistrantID"
                w_Sql = w_Sql & " ) Values ("
                w_Sql = w_Sql & " '" & General.g_strHospitalCD & "'"
                w_Sql = w_Sql & ", " & Convert.ToString(p_intCurrentDate)
                w_Sql = w_Sql & ", '" & p_strStaffDataID & "'"
                w_Sql = w_Sql & ", '" & p_strKinmuCD & "'"
                w_Sql = w_Sql & ", '" & p_strRiyuKBN & "'"
                w_Sql = w_Sql & ", '" & p_strComment & "'" '2015/04/10 Bando Add 
                w_Sql = w_Sql & ", " & p_strSysDate
                w_Sql = w_Sql & ", " & p_strSysDate
                w_Sql = w_Sql & ", '" & General.g_strUserID & "'"
                w_Sql = w_Sql & " )"
                w_Sql = w_Sql & " "
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "Insert Into NS_KINMUPLAN_F ("
                w_Sql = w_Sql & " HospitalCD"
                w_Sql = w_Sql & ", DateF"
                w_Sql = w_Sql & ", StaffMngID"
                w_Sql = w_Sql & ", KinmuCD"
                w_Sql = w_Sql & ", ReasonKbn"
                w_Sql = w_Sql & ", HOPECOMMENT" '2015/04/10 Bando Add
                w_Sql = w_Sql & ", RegistFirstTimeDate"
                w_Sql = w_Sql & ", LastUpdTimeDate"
                w_Sql = w_Sql & ", RegistrantID"
                w_Sql = w_Sql & " ) Values ("
                w_Sql = w_Sql & " '" & General.g_strHospitalCD & "'"
                w_Sql = w_Sql & ", " & Convert.ToString(p_intCurrentDate)
                w_Sql = w_Sql & ", '" & p_strStaffDataID & "'"
                w_Sql = w_Sql & ", '" & p_strKinmuCD & "'"
                w_Sql = w_Sql & ", '" & p_strRiyuKBN & "'"
                w_Sql = w_Sql & ", '" & p_strComment & "'" '2015/04/10 Bando Add 
                w_Sql = w_Sql & ", " & p_strSysDate
                w_Sql = w_Sql & ", " & p_strSysDate
                w_Sql = w_Sql & ", '" & General.g_strUserID & "'"
                w_Sql = w_Sql & " )"
                w_Sql = w_Sql & " "
            End If

            Call General.paDBExecute(w_Sql)
            insert_NS_KINMUPLAN_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="p_intCurrentDate"></param>
    ''' <param name="p_strStaffID"></param>
    ''' <param name="p_strKinmuCD"></param>
    ''' <param name="p_strRiyuKBN"></param>
    ''' <param name="p_strKangoCD"></param>
    ''' <param name="p_strComment"></param>
    ''' <param name="p_strSysDate"></param>
    ''' <returns></returns>
    Public Function insert_NS_KINMUPLAN_F_02(ByVal p_intCurrentDate As Integer,
                                             ByVal p_strStaffID As String,
                                             ByVal p_strKinmuCD As String,
                                             ByVal p_strRiyuKBN As String,
                                             ByVal p_strKangoCD As String,
                                             ByVal p_strComment As String,
                                             ByVal p_strSysDate As String) As Boolean
        insert_NS_KINMUPLAN_F_02 = False

        Try
            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "INSERT INTO NS_KINMUPLAN_F "
                w_Sql = w_Sql & "(HOSPITALCD,DATEF,STAFFMNGID,KINMUCD,REASONKBN,OUENKINMUDEPTCD,HOPECOMMENT,REGISTFIRSTTIMEDATE,LASTUPDTIMEDATE,REGISTRANTID)"
                w_Sql = w_Sql & "VALUES('" & General.g_strHospitalCD & "'," & p_intCurrentDate & ",'" & p_strStaffID & "','" & p_strKinmuCD
                w_Sql = w_Sql & "','" & p_strRiyuKBN & "','" & p_strKangoCD & "','" & p_strComment & "'," & p_strSysDate & "," & p_strSysDate & ",'" & General.g_strUserID & "')"
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "INSERT INTO NS_KINMUPLAN_F "
                w_Sql = w_Sql & "(HOSPITALCD,DATEF,STAFFMNGID,KINMUCD,REASONKBN,OUENKINMUDEPTCD,HOPECOMMENT,REGISTFIRSTTIMEDATE,LASTUPDTIMEDATE,REGISTRANTID)"
                w_Sql = w_Sql & "VALUES('" & General.g_strHospitalCD & "'," & p_intCurrentDate & ",'" & p_strStaffID & "','" & p_strKinmuCD
                w_Sql = w_Sql & "','" & p_strRiyuKBN & "','" & p_strKangoCD & "','" & p_strComment & "'," & p_strSysDate & "," & p_strSysDate & ",'" & General.g_strUserID & "')"
            End If

            Call General.paDBExecute(w_Sql)
            insert_NS_KINMUPLAN_F_02 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="p_intCurrentDate"></param>
    ''' <param name="p_strStaffID"></param>
    ''' <param name="p_strKinmuCD"></param>
    ''' <param name="p_strRiyuKBN"></param>
    ''' <param name="p_strKangoCD"></param>
    ''' <param name="p_strComment"></param>
    ''' <param name="p_strSysDate"></param>
    ''' <returns></returns>
    Public Function insert_NS_KINMURESULT_F_01(ByVal p_intCurrentDate As Integer,
                                               ByVal p_strStaffID As String,
                                               ByVal p_strKinmuCD As String,
                                               ByVal p_strRiyuKBN As String,
                                               ByVal p_strKangoCD As String,
                                               ByVal p_strComment As String,
                                               ByVal p_strSysDate As String) As Boolean
        insert_NS_KINMURESULT_F_01 = False

        Try

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "Insert Into NS_KINMURESULT_F ("
                w_Sql = w_Sql & " HospitalCD"
                w_Sql = w_Sql & ", DateF"
                w_Sql = w_Sql & ", StaffMngID"
                w_Sql = w_Sql & ", KinmuCD"
                w_Sql = w_Sql & ", ReasonKbn"
                w_Sql = w_Sql & ", OUENKINMUDEPTCD"
                w_Sql = w_Sql & ", DecDeptCD"
                w_Sql = w_Sql & ", HOPECOMMENT" '2015/04/10 Bando Add
                w_Sql = w_Sql & ", RegistFirstTimeDate"
                w_Sql = w_Sql & ", LastUpdTimeDate"
                w_Sql = w_Sql & ", RegistrantID"
                w_Sql = w_Sql & " ) Values ("
                w_Sql = w_Sql & " '" & General.g_strHospitalCD & "'" 'HospitalCD
                w_Sql = w_Sql & ", " & p_intCurrentDate '年月日
                w_Sql = w_Sql & ", '" & p_strStaffID & "'" '職員管理番号
                w_Sql = w_Sql & ", '" & p_strKinmuCD & "'" 'KinmuCD
                w_Sql = w_Sql & ", '" & p_strRiyuKBN & "'" '理由区分
                w_Sql = w_Sql & ", '" & p_strKangoCD & "'" '応援先看護単位CD
                w_Sql = w_Sql & ", '" & General.g_strSelKinmuDeptCD & "'" '確定部署CD（看護単位CD）
                w_Sql = w_Sql & ", '" & p_strComment & "'" '希望コメント　2015/04/10 Bando Add
                w_Sql = w_Sql & ", " & p_strSysDate 'RegistFirstTimeDate
                w_Sql = w_Sql & ", " & p_strSysDate 'LastUpdTimeDate
                w_Sql = w_Sql & ", '" & General.g_strUserID & "'" 'RegistrantID
                w_Sql = w_Sql & " ) "
                w_Sql = w_Sql & " "

            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "Insert Into NS_KINMURESULT_F ("
                w_Sql = w_Sql & " HospitalCD"
                w_Sql = w_Sql & ", DateF"
                w_Sql = w_Sql & ", StaffMngID"
                w_Sql = w_Sql & ", KinmuCD"
                w_Sql = w_Sql & ", ReasonKbn"
                w_Sql = w_Sql & ", OUENKINMUDEPTCD"
                w_Sql = w_Sql & ", DecDeptCD"
                w_Sql = w_Sql & ", HOPECOMMENT" '2015/04/10 Bando Add
                w_Sql = w_Sql & ", RegistFirstTimeDate"
                w_Sql = w_Sql & ", LastUpdTimeDate"
                w_Sql = w_Sql & ", RegistrantID"
                w_Sql = w_Sql & " ) Values ("
                w_Sql = w_Sql & " '" & General.g_strHospitalCD & "'" 'HospitalCD
                w_Sql = w_Sql & ", " & p_intCurrentDate '年月日
                w_Sql = w_Sql & ", '" & p_strStaffID & "'" '職員管理番号
                w_Sql = w_Sql & ", '" & p_strKinmuCD & "'" 'KinmuCD
                w_Sql = w_Sql & ", '" & p_strRiyuKBN & "'" '理由区分
                w_Sql = w_Sql & ", '" & p_strKangoCD & "'" '応援先看護単位CD
                w_Sql = w_Sql & ", '" & General.g_strSelKinmuDeptCD & "'" '確定部署CD（看護単位CD）
                w_Sql = w_Sql & ", '" & p_strComment & "'" '希望コメント　2015/04/10 Bando Add
                w_Sql = w_Sql & ", " & p_strSysDate 'RegistFirstTimeDate
                w_Sql = w_Sql & ", " & p_strSysDate 'LastUpdTimeDate
                w_Sql = w_Sql & ", '" & General.g_strUserID & "'" 'RegistrantID
                w_Sql = w_Sql & " ) "
                w_Sql = w_Sql & " "

            End If
            Call General.paDBExecute(w_Sql)
            insert_NS_KINMURESULT_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="p_strStaffID"></param>
    ''' <param name="p_objStaffData"></param>
    ''' <param name="p_strDate"></param>
    ''' <param name="p_intSEQIdx"></param>
    ''' <param name="p_intIndex"></param>
    ''' <param name="p_intDateIdx"></param>
    ''' <param name="p_strSysDate"></param>
    ''' <returns></returns>
    Public Function insert_NS_NENKYU_F_01(ByVal p_strStaffID As String,
                                          ByVal p_objStaffData As Object,
                                          ByVal p_strDate As String,
                                          ByVal p_intSEQIdx As Integer,
                                          ByVal p_intIndex As Integer,
                                          ByVal p_intDateIdx As Integer,
                                          ByVal p_strSysDate As String) As Boolean
        '初期化
        insert_NS_NENKYU_F_01 = False

        Try

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "Insert Into NS_NENKYU_F ("
                w_Sql = w_Sql & "HospitalCD,"
                w_Sql = w_Sql & "StaffMngID,"
                w_Sql = w_Sql & "DateF,"
                w_Sql = w_Sql & "SEQ,"
                w_Sql = w_Sql & "GETCONTENTSKBN,"
                w_Sql = w_Sql & "HolidayBunruiCD,"
                w_Sql = w_Sql & "FromTime,"
                w_Sql = w_Sql & "ToTime,"
                w_Sql = w_Sql & "NEXTDAYFLG,"
                w_Sql = w_Sql & "KINMUDATE,"
                w_Sql = w_Sql & "DateKbn,"
                w_Sql = w_Sql & "UNIQUESEQNO,"
                w_Sql = w_Sql & "APPROVEFLG,"
                w_Sql = w_Sql & "DELFLG,"
                w_Sql = w_Sql & "NENKYUTIME,"
                w_Sql = w_Sql & "HOLSUBFLG,"
                w_Sql = w_Sql & "DAYTIME,"
                w_Sql = w_Sql & "NIGHTTIME,"
                w_Sql = w_Sql & "NEXTNIGHTTIME,"
                w_Sql = w_Sql & "RegistFirstTimeDate,"
                w_Sql = w_Sql & "LastUpdTimeDate,"
                w_Sql = w_Sql & "RegistrantID)"
                w_Sql = w_Sql & "Values('"
                w_Sql = w_Sql & General.g_strHospitalCD & "'," '病院CD
                w_Sql = w_Sql & "'" & p_strStaffID & "'," '職員管理番号
                w_Sql = w_Sql & Convert.ToString(p_strDate) & "," '日付
                w_Sql = w_Sql & Convert.ToString(p_intSEQIdx) & "," 'SEQ
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).NenkyuData(p_intDateIdx).Detail(p_intSEQIdx).GetContentsKbn & "'," '取得内容区分
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).NenkyuData(p_intDateIdx).Detail(p_intSEQIdx).HolidayBunruiCD & "'," '休暇分類CD
                w_Sql = w_Sql & Convert.ToString(p_objStaffData(p_intIndex).NenkyuData(p_intDateIdx).Detail(p_intSEQIdx).FromTime) & "," '開始時間
                w_Sql = w_Sql & Convert.ToString(p_objStaffData(p_intIndex).NenkyuData(p_intDateIdx).Detail(p_intSEQIdx).ToTime) & "," '終了時間
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).NenkyuData(p_intDateIdx).Detail(p_intSEQIdx).DateKbn & "'," '翌日FLG
                w_Sql = w_Sql & Convert.ToString(p_strDate) & "," '勤務年月日
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).NenkyuData(p_intDateIdx).Detail(p_intSEQIdx).DateKbn & "'," '年月日区分
                w_Sql = w_Sql & "''," 'UNIQUESEQNO
                w_Sql = w_Sql & "'1'," '承認済FLG
                w_Sql = w_Sql & "''," '削除FLG
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).NenkyuData(p_intDateIdx).Detail(p_intSEQIdx).NenkyuTime & "'," '時間年休
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).NenkyuData(p_intDateIdx).Detail(p_intSEQIdx).HolSubFlg & "'," '休憩減算フラグ
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).NenkyuData(p_intDateIdx).Detail(p_intSEQIdx).DayTime & "'," '日勤時間
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).NenkyuData(p_intDateIdx).Detail(p_intSEQIdx).NightTime & "'," '夜勤時間
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).NenkyuData(p_intDateIdx).Detail(p_intSEQIdx).NextNightTime & "'," '翌日夜勤時間
                w_Sql = w_Sql & p_strSysDate & ","
                w_Sql = w_Sql & p_strSysDate & ","
                w_Sql = w_Sql & "'" & General.g_strUserID & "')"
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "Insert Into NS_NENKYU_F ("
                w_Sql = w_Sql & "HospitalCD,"
                w_Sql = w_Sql & "StaffMngID,"
                w_Sql = w_Sql & "DateF,"
                w_Sql = w_Sql & "SEQ,"
                w_Sql = w_Sql & "GETCONTENTSKBN,"
                w_Sql = w_Sql & "HolidayBunruiCD,"
                w_Sql = w_Sql & "FromTime,"
                w_Sql = w_Sql & "ToTime,"
                w_Sql = w_Sql & "NEXTDAYFLG,"
                w_Sql = w_Sql & "KINMUDATE,"
                w_Sql = w_Sql & "DateKbn,"
                w_Sql = w_Sql & "UNIQUESEQNO,"
                w_Sql = w_Sql & "APPROVEFLG,"
                w_Sql = w_Sql & "DELFLG,"
                w_Sql = w_Sql & "NENKYUTIME,"
                w_Sql = w_Sql & "HOLSUBFLG,"
                w_Sql = w_Sql & "DAYTIME,"
                w_Sql = w_Sql & "NIGHTTIME,"
                w_Sql = w_Sql & "NEXTNIGHTTIME,"
                w_Sql = w_Sql & "RegistFirstTimeDate,"
                w_Sql = w_Sql & "LastUpdTimeDate,"
                w_Sql = w_Sql & "RegistrantID)"
                w_Sql = w_Sql & "Values('"
                w_Sql = w_Sql & General.g_strHospitalCD & "'," '病院CD
                w_Sql = w_Sql & "'" & p_strStaffID & "'," '職員管理番号
                w_Sql = w_Sql & Convert.ToString(p_strDate) & "," '日付
                w_Sql = w_Sql & Convert.ToString(p_intSEQIdx) & "," 'SEQ
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).NenkyuData(p_intDateIdx).Detail(p_intSEQIdx).GetContentsKbn & "'," '取得内容区分
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).NenkyuData(p_intDateIdx).Detail(p_intSEQIdx).HolidayBunruiCD & "'," '休暇分類CD
                w_Sql = w_Sql & Convert.ToString(p_objStaffData(p_intIndex).NenkyuData(p_intDateIdx).Detail(p_intSEQIdx).FromTime) & "," '開始時間
                w_Sql = w_Sql & Convert.ToString(p_objStaffData(p_intIndex).NenkyuData(p_intDateIdx).Detail(p_intSEQIdx).ToTime) & "," '終了時間
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).NenkyuData(p_intDateIdx).Detail(p_intSEQIdx).DateKbn & "'," '翌日FLG
                w_Sql = w_Sql & Convert.ToString(p_strDate) & "," '勤務年月日
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).NenkyuData(p_intDateIdx).Detail(p_intSEQIdx).DateKbn & "'," '年月日区分
                w_Sql = w_Sql & "''," 'UNIQUESEQNO
                w_Sql = w_Sql & "'1'," '承認済FLG
                w_Sql = w_Sql & "''," '削除FLG
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).NenkyuData(p_intDateIdx).Detail(p_intSEQIdx).NenkyuTime & "'," '時間年休
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).NenkyuData(p_intDateIdx).Detail(p_intSEQIdx).HolSubFlg & "'," '休憩減算フラグ
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).NenkyuData(p_intDateIdx).Detail(p_intSEQIdx).DayTime & "'," '日勤時間
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).NenkyuData(p_intDateIdx).Detail(p_intSEQIdx).NightTime & "'," '夜勤時間
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).NenkyuData(p_intDateIdx).Detail(p_intSEQIdx).NextNightTime & "'," '翌日夜勤時間
                w_Sql = w_Sql & p_strSysDate & ","
                w_Sql = w_Sql & p_strSysDate & ","
                w_Sql = w_Sql & "'" & General.g_strUserID & "')"
            End If

            Call General.paDBExecute(w_Sql)
            insert_NS_NENKYU_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="p_intCurrentDate"></param>
    ''' <param name="p_strStaffID"></param>
    ''' <param name="p_intSeq"></param>
    ''' <param name="p_strKinmuCD_Old"></param>
    ''' <param name="p_strKinmuCD_New"></param>
    ''' <param name="p_intSysDate"></param>
    ''' <returns></returns>
    Public Function insert_NS_KINMUCHGHISTORY_F_01(ByVal p_intCurrentDate As Integer,
                                                   ByVal p_strStaffID As String,
                                                   ByVal p_intSeq As Integer,
                                                   ByVal p_strKinmuCD_Old As String,
                                                   ByVal p_strKinmuCD_New As String,
                                                   ByVal p_intSysDate As Integer) As Boolean
        insert_NS_KINMUCHGHISTORY_F_01 = False

        Try
            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "Insert Into NS_KINMUCHGHISTORY_F "
                w_Sql = w_Sql & "(HospitalCD,DateF,StaffMngID,SEQ,ChgBeforeKinmu,ChgAfterKinmu,LastUpdTimeDate,RegistrantID)"
                w_Sql = w_Sql & "Values('" & General.g_strHospitalCD & "'," & p_intCurrentDate & ",'" & p_strStaffID & "'," & p_intSeq & ",'" & p_strKinmuCD_Old
                w_Sql = w_Sql & "','" & p_strKinmuCD_New & "'," & p_intSysDate & ",'" & General.g_strUserID & "')"

            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "Insert Into NS_KINMUCHGHISTORY_F "
                w_Sql = w_Sql & "(HospitalCD,DateF,StaffMngID,SEQ,ChgBeforeKinmu,ChgAfterKinmu,LastUpdTimeDate,RegistrantID)"
                w_Sql = w_Sql & "Values('" & General.g_strHospitalCD & "'," & p_intCurrentDate & ",'" & p_strStaffID & "'," & p_intSeq & ",'" & p_strKinmuCD_Old
                w_Sql = w_Sql & "','" & p_strKinmuCD_New & "'," & p_intSysDate & ",'" & General.g_strUserID & "')"

            End If

            Call General.paDBExecute(w_Sql)
            insert_NS_KINMUCHGHISTORY_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="p_intPlanNo"></param>
    ''' <param name="p_strKakuteiDate"></param>
    ''' <param name="p_strSysDate"></param>
    ''' <returns></returns>
    Public Function insert_NS_PLANDECISION_F_01(ByVal p_intPlanNo As Integer,
                                                ByVal p_strKakuteiDate As String,
                                                ByVal p_strSysDate As String) As Boolean
        insert_NS_PLANDECISION_F_01 = False

        Try

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "INSERT INTO NS_PLANDECISION_F "
                w_Sql = w_Sql & "(HOSPITALCD,KINMUDEPTCD,PLANNO,DECISIONDATE,LASTUPDTIMEDATE,REGISTRANTID)"
                w_Sql = w_Sql & "VALUES('" & General.g_strHospitalCD & "','" & General.g_strSelKinmuDeptCD & "'," & p_intPlanNo & "," & p_strKakuteiDate
                w_Sql = w_Sql & "," & p_strSysDate & ",'" & General.g_strUserID & "')"
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "INSERT INTO NS_PLANDECISION_F "
                w_Sql = w_Sql & "(HOSPITALCD,KINMUDEPTCD,PLANNO,DECISIONDATE,LASTUPDTIMEDATE,REGISTRANTID)"
                w_Sql = w_Sql & "VALUES('" & General.g_strHospitalCD & "','" & General.g_strSelKinmuDeptCD & "'," & p_intPlanNo & "," & p_strKakuteiDate
                w_Sql = w_Sql & "," & p_strSysDate & ",'" & General.g_strUserID & "')"
            End If
            Call General.paDBExecute(w_Sql)
            insert_NS_PLANDECISION_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="p_strPlanNo"></param>
    ''' <param name="p_objStaffData"></param>
    ''' <param name="p_intIndex"></param>
    ''' <param name="p_intDispNo1"></param>
    ''' <param name="p_intDispNo2"></param>
    ''' <param name="p_intDispNo3"></param>
    ''' <param name="p_intDispNo4"></param>
    ''' <param name="p_intDispNo5"></param>
    ''' <param name="p_dblSysDate"></param>
    ''' <returns></returns>
    Public Function insert_NS_STAFFHISTORY_F_01(ByVal p_strPlanNo As String,
                                                ByVal p_objStaffData As Object,
                                                ByVal p_intIndex As Integer,
                                                ByVal p_intDispNo1 As Integer,
                                                ByVal p_intDispNo2 As Integer,
                                                ByVal p_intDispNo3 As Integer,
                                                ByVal p_intDispNo4 As Integer,
                                                ByVal p_intDispNo5 As Integer,
                                                ByVal p_dblSysDate As Double) As Boolean
        insert_NS_STAFFHISTORY_F_01 = False

        Try

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "INSERT INTO NS_STAFFHISTORY_F ( "
                w_Sql = w_Sql & "HOSPITALCD, "
                w_Sql = w_Sql & "PLANNO, "
                w_Sql = w_Sql & "KINMUDEPTCD, "
                w_Sql = w_Sql & "STAFFMNGID, "
                w_Sql = w_Sql & "SKILLLVLCD, "
                w_Sql = w_Sql & "DISPNO, "
                w_Sql = w_Sql & "DISPNO1, "
                w_Sql = w_Sql & "DISPNO2, "
                w_Sql = w_Sql & "DISPNO3, "
                w_Sql = w_Sql & "DISPNO4, "
                w_Sql = w_Sql & "DISPNO5, "
                w_Sql = w_Sql & "AUTOALLOCKBN, "
                w_Sql = w_Sql & "TEAM, "
                w_Sql = w_Sql & "NIGHTONLYSTAFFKBN, "
                w_Sql = w_Sql & "PATTERNCD, "
                w_Sql = w_Sql & "REGISTFIRSTTIMEDATE, "
                w_Sql = w_Sql & "LASTUPDTIMEDATE, "
                w_Sql = w_Sql & "REGISTRANTID "
                w_Sql = w_Sql & ") VALUES ( "
                w_Sql = w_Sql & "'" & General.g_strHospitalCD & "', "
                w_Sql = w_Sql & p_strPlanNo & ", "
                w_Sql = w_Sql & "'" & General.g_strSelKinmuDeptCD & "', "
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).ID & "', "
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).GiryoLvCD & "', "
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).HyojiNo & "', "
                w_Sql = w_Sql & "'" & p_intDispNo1 & "', "
                w_Sql = w_Sql & "'" & p_intDispNo2 & "', "
                w_Sql = w_Sql & "'" & p_intDispNo3 & "', "
                w_Sql = w_Sql & "'" & p_intDispNo4 & "', "
                w_Sql = w_Sql & "'" & p_intDispNo5 & "', "
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).AutoKBN & "', "
                w_Sql = w_Sql & p_objStaffData(p_intIndex).Team & ", "
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).YakinKBN & "', "
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).PatternCD & "', "
                w_Sql = w_Sql & p_dblSysDate & ", "
                w_Sql = w_Sql & p_dblSysDate & ", "
                w_Sql = w_Sql & "'" & General.g_strUserID & "' "
                w_Sql = w_Sql & ") "
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "INSERT INTO NS_STAFFHISTORY_F ( "
                w_Sql = w_Sql & "HOSPITALCD, "
                w_Sql = w_Sql & "PLANNO, "
                w_Sql = w_Sql & "KINMUDEPTCD, "
                w_Sql = w_Sql & "STAFFMNGID, "
                w_Sql = w_Sql & "SKILLLVLCD, "
                w_Sql = w_Sql & "DISPNO, "
                w_Sql = w_Sql & "DISPNO1, "
                w_Sql = w_Sql & "DISPNO2, "
                w_Sql = w_Sql & "DISPNO3, "
                w_Sql = w_Sql & "DISPNO4, "
                w_Sql = w_Sql & "DISPNO5, "
                w_Sql = w_Sql & "AUTOALLOCKBN, "
                w_Sql = w_Sql & "TEAM, "
                w_Sql = w_Sql & "NIGHTONLYSTAFFKBN, "
                w_Sql = w_Sql & "PATTERNCD, "
                w_Sql = w_Sql & "REGISTFIRSTTIMEDATE, "
                w_Sql = w_Sql & "LASTUPDTIMEDATE, "
                w_Sql = w_Sql & "REGISTRANTID "
                w_Sql = w_Sql & ") VALUES ( "
                w_Sql = w_Sql & "'" & General.g_strHospitalCD & "', "
                w_Sql = w_Sql & p_strPlanNo & ", "
                w_Sql = w_Sql & "'" & General.g_strSelKinmuDeptCD & "', "
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).ID & "', "
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).GiryoLvCD & "', "
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).HyojiNo & "', "
                w_Sql = w_Sql & "'" & p_intDispNo1 & "', "
                w_Sql = w_Sql & "'" & p_intDispNo2 & "', "
                w_Sql = w_Sql & "'" & p_intDispNo3 & "', "
                w_Sql = w_Sql & "'" & p_intDispNo4 & "', "
                w_Sql = w_Sql & "'" & p_intDispNo5 & "', "
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).AutoKBN & "', "
                w_Sql = w_Sql & p_objStaffData(p_intIndex).Team & ", "
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).YakinKBN & "', "
                w_Sql = w_Sql & "'" & p_objStaffData(p_intIndex).PatternCD & "', "
                w_Sql = w_Sql & p_dblSysDate & ", "
                w_Sql = w_Sql & p_dblSysDate & ", "
                w_Sql = w_Sql & "'" & General.g_strUserID & "' "
                w_Sql = w_Sql & ") "
            End If
            Call General.paDBExecute(w_Sql)
            insert_NS_STAFFHISTORY_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="p_objStaffData"></param>
    ''' <param name="p_intIndex"></param>
    ''' <param name="p_intInt"></param>
    ''' <param name="p_dblSysDate"></param>
    ''' <returns></returns>
    Public Function insert_NS_DAIKYUMNG_F_01(ByVal p_objStaffData As Object,
                                             ByVal p_intIndex As Integer,
                                             ByVal p_intInt As Integer,
                                             ByVal p_dblSysDate As Double) As Boolean
        insert_NS_DAIKYUMNG_F_01 = False

        Try
            w_sqlBuilder = New System.Text.StringBuilder

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "Insert Into NS_DAIKYUMNG_F "
                w_Sql = w_Sql & "(HospitalCD,"
                w_Sql = w_Sql & " StaffMngID,"
                w_Sql = w_Sql & " GetKbn,"
                w_Sql = w_Sql & " WorkHolKinmuDate,"
                w_Sql = w_Sql & " WorkHolKinmuCD,"
                w_Sql = w_Sql & " TodokedeNo,"
                w_Sql = w_Sql & " RegistFirstTimeDate,"
                w_Sql = w_Sql & " LastUpdTimeDate,"
                w_Sql = w_Sql & " RegistrantID) "
                w_Sql = w_Sql & " Values('" & General.g_strHospitalCD & "','"
                w_Sql = w_Sql & p_objStaffData(p_intIndex).ID & "','"
                w_Sql = w_Sql & p_objStaffData(p_intIndex).Daikyu(p_intInt).GetKbn & "',"
                w_Sql = w_Sql & p_objStaffData(p_intIndex).Daikyu(p_intInt).HolDate & ",'"
                w_Sql = w_Sql & p_objStaffData(p_intIndex).Daikyu(p_intInt).HolKinmuCD & "',"
                w_Sql = w_Sql & "0,"
                w_Sql = w_Sql & p_dblSysDate & ","
                w_Sql = w_Sql & p_dblSysDate & ",'"
                w_Sql = w_Sql & General.g_strUserID & "')"
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "Insert Into NS_DAIKYUMNG_F "
                w_Sql = w_Sql & "(HospitalCD,"
                w_Sql = w_Sql & " StaffMngID,"
                w_Sql = w_Sql & " GetKbn,"
                w_Sql = w_Sql & " WorkHolKinmuDate,"
                w_Sql = w_Sql & " WorkHolKinmuCD,"
                w_Sql = w_Sql & " TodokedeNo,"
                w_Sql = w_Sql & " RegistFirstTimeDate,"
                w_Sql = w_Sql & " LastUpdTimeDate,"
                w_Sql = w_Sql & " RegistrantID) "
                w_Sql = w_Sql & " Values('" & General.g_strHospitalCD & "','"
                w_Sql = w_Sql & p_objStaffData(p_intIndex).ID & "','"
                w_Sql = w_Sql & p_objStaffData(p_intIndex).Daikyu(p_intInt).GetKbn & "',"
                w_Sql = w_Sql & p_objStaffData(p_intIndex).Daikyu(p_intInt).HolDate & ",'"
                w_Sql = w_Sql & p_objStaffData(p_intIndex).Daikyu(p_intInt).HolKinmuCD & "',"
                w_Sql = w_Sql & "0,"
                w_Sql = w_Sql & p_dblSysDate & ","
                w_Sql = w_Sql & p_dblSysDate & ",'"
                w_Sql = w_Sql & General.g_strUserID & "')"
            End If


            Call General.paDBExecute(w_Sql)
            insert_NS_DAIKYUMNG_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="p_objStaffData"></param>
    ''' <param name="p_intIndex"></param>
    ''' <param name="p_intInt"></param>
    ''' <param name="p_intlngLoop"></param>
    ''' <param name="p_dblSysDate"></param>
    ''' <returns></returns>
    Public Function insert_NS_DAIKYUDETAILMNG_F_01(ByVal p_objStaffData As Object,
                                                   ByVal p_intIndex As Integer,
                                                   ByVal p_intInt As Integer,
                                                   ByVal p_intlngLoop As Integer,
                                                   ByVal p_dblSysDate As Double) As Boolean
        insert_NS_DAIKYUDETAILMNG_F_01 = False

        Try
            w_sqlBuilder = New System.Text.StringBuilder

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "Insert Into NS_DAIKYUDETAILMNG_F "
                w_Sql = w_Sql & "(HospitalCD,"
                w_Sql = w_Sql & " StaffMngID,"
                w_Sql = w_Sql & " WorkHolKinmuDate,"
                w_Sql = w_Sql & " SEQ,"
                w_Sql = w_Sql & " GetFlg,"
                w_Sql = w_Sql & " GetDaikyuDate,"
                w_Sql = w_Sql & " GetDaikyuKinmuCD,"
                w_Sql = w_Sql & " RegistFirstTimeDate,"
                w_Sql = w_Sql & " LastUpdTimeDate,"
                w_Sql = w_Sql & " RegistrantID) "
                w_Sql = w_Sql & " Values('" & General.g_strHospitalCD & "','"
                w_Sql = w_Sql & p_objStaffData(p_intIndex).ID & "',"
                w_Sql = w_Sql & p_objStaffData(p_intIndex).Daikyu(p_intInt).HolDate & ","
                w_Sql = w_Sql & p_intlngLoop & ",'"
                w_Sql = w_Sql & p_objStaffData(p_intIndex).Daikyu(p_intInt).DaikyuDetail(p_intlngLoop).GetFlg & "',"
                w_Sql = w_Sql & p_objStaffData(p_intIndex).Daikyu(p_intInt).DaikyuDetail(p_intlngLoop).DaikyuDate & ",'"
                w_Sql = w_Sql & p_objStaffData(p_intIndex).Daikyu(p_intInt).DaikyuDetail(p_intlngLoop).DaikyuKinmuCD & "',"
                w_Sql = w_Sql & p_dblSysDate & ","
                w_Sql = w_Sql & p_dblSysDate & ",'"
                w_Sql = w_Sql & General.g_strUserID & "')"
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "Insert Into NS_DAIKYUDETAILMNG_F "
                w_Sql = w_Sql & "(HospitalCD,"
                w_Sql = w_Sql & " StaffMngID,"
                w_Sql = w_Sql & " WorkHolKinmuDate,"
                w_Sql = w_Sql & " SEQ,"
                w_Sql = w_Sql & " GetFlg,"
                w_Sql = w_Sql & " GetDaikyuDate,"
                w_Sql = w_Sql & " GetDaikyuKinmuCD,"
                w_Sql = w_Sql & " RegistFirstTimeDate,"
                w_Sql = w_Sql & " LastUpdTimeDate,"
                w_Sql = w_Sql & " RegistrantID) "
                w_Sql = w_Sql & " Values('" & General.g_strHospitalCD & "','"
                w_Sql = w_Sql & p_objStaffData(p_intIndex).ID & "',"
                w_Sql = w_Sql & p_objStaffData(p_intIndex).Daikyu(p_intInt).HolDate & ","
                w_Sql = w_Sql & p_intlngLoop & ",'"
                w_Sql = w_Sql & p_objStaffData(p_intIndex).Daikyu(p_intInt).DaikyuDetail(p_intlngLoop).GetFlg & "',"
                w_Sql = w_Sql & p_objStaffData(p_intIndex).Daikyu(p_intInt).DaikyuDetail(p_intlngLoop).DaikyuDate & ",'"
                w_Sql = w_Sql & p_objStaffData(p_intIndex).Daikyu(p_intInt).DaikyuDetail(p_intlngLoop).DaikyuKinmuCD & "',"
                w_Sql = w_Sql & p_dblSysDate & ","
                w_Sql = w_Sql & p_dblSysDate & ",'"
                w_Sql = w_Sql & General.g_strUserID & "')"
            End If

            Call General.paDBExecute(w_Sql)
            insert_NS_DAIKYUDETAILMNG_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HJ
    ''' </summary>
    ''' <param name="p_strMngStaffID"></param>
    ''' <param name="p_intSelDate"></param>
    ''' <param name="p_intDataLoop"></param>
    ''' <param name="p_strGetDaikyuType"></param>
    ''' <param name="p_intlngYMD"></param>
    ''' <param name="p_strKinmuCD"></param>
    ''' <param name="p_dblRegistFirstTimeDate"></param>
    ''' <param name="p_dblSysDate"></param>
    ''' <returns></returns>
    Public Function insert_NS_DAIKYUDETAILMNG_F_02(ByVal p_strMngStaffID As String,
                                                   ByVal p_intSelDate As Integer,
                                                   ByVal p_intDataLoop As Integer,
                                                   ByVal p_strGetDaikyuType As String,
                                                   ByVal p_intlngYMD As Integer,
                                                   ByVal p_strKinmuCD As String,
                                                   ByVal p_dblRegistFirstTimeDate As Double,
                                                   ByVal p_dblSysDate As Double) As Boolean
        insert_NS_DAIKYUDETAILMNG_F_02 = False

        Try
            w_sqlBuilder = New System.Text.StringBuilder

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "Insert Into NS_DAIKYUDETAILMNG_F ("
                w_Sql = w_Sql & "HospitalCD,"
                w_Sql = w_Sql & "StaffMngID,"
                w_Sql = w_Sql & "WorkHolKinmuDate,"
                w_Sql = w_Sql & "SEQ,"
                w_Sql = w_Sql & "GETFLG,"
                w_Sql = w_Sql & "GetDaikyuDate,"
                w_Sql = w_Sql & "GetDaikyuKinmuCD,"
                w_Sql = w_Sql & "RegistFirstTimeDate,"
                w_Sql = w_Sql & "LastUpdTimeDate,"
                w_Sql = w_Sql & "RegistrantID)"
                w_Sql = w_Sql & "Values('"
                w_Sql = w_Sql & Trim(General.g_strHospitalCD) & "'," '病院CD
                w_Sql = w_Sql & "'" & Trim(p_strMngStaffID) & "'," '職員管理番号
                w_Sql = w_Sql & p_intSelDate & "," '発生日
                w_Sql = w_Sql & p_intDataLoop & "," 'SEQ
                w_Sql = w_Sql & "'" & Trim(p_strGetDaikyuType) & "'," '取得タイプ
                w_Sql = w_Sql & p_intlngYMD & "," '取得日
                w_Sql = w_Sql & "'" & Trim(p_strKinmuCD) & "'," '取得勤務CD
                w_Sql = w_Sql & p_dblRegistFirstTimeDate & ","
                w_Sql = w_Sql & p_dblSysDate & ","
                w_Sql = w_Sql & "'" & General.g_strUserID & "')"
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "Insert Into NS_DAIKYUDETAILMNG_F ("
                w_Sql = w_Sql & "HospitalCD,"
                w_Sql = w_Sql & "StaffMngID,"
                w_Sql = w_Sql & "WorkHolKinmuDate,"
                w_Sql = w_Sql & "SEQ,"
                w_Sql = w_Sql & "GETFLG,"
                w_Sql = w_Sql & "GetDaikyuDate,"
                w_Sql = w_Sql & "GetDaikyuKinmuCD,"
                w_Sql = w_Sql & "RegistFirstTimeDate,"
                w_Sql = w_Sql & "LastUpdTimeDate,"
                w_Sql = w_Sql & "RegistrantID)"
                w_Sql = w_Sql & "Values('"
                w_Sql = w_Sql & Trim(General.g_strHospitalCD) & "'," '病院CD
                w_Sql = w_Sql & "'" & Trim(p_strMngStaffID) & "'," '職員管理番号
                w_Sql = w_Sql & p_intSelDate & "," '発生日
                w_Sql = w_Sql & p_intDataLoop & "," 'SEQ
                w_Sql = w_Sql & "'" & Trim(p_strGetDaikyuType) & "'," '取得タイプ
                w_Sql = w_Sql & p_intlngYMD & "," '取得日
                w_Sql = w_Sql & "'" & Trim(p_strKinmuCD) & "'," '取得勤務CD
                w_Sql = w_Sql & p_dblRegistFirstTimeDate & ","
                w_Sql = w_Sql & p_dblSysDate & ","
                w_Sql = w_Sql & "'" & General.g_strUserID & "')"
            End If

            Call General.paDBExecute(w_Sql)
            insert_NS_DAIKYUDETAILMNG_F_02 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="p_strKinmuCD"></param>
    ''' <param name="p_strRiyuKBN"></param>
    ''' <param name="p_strComment"></param>
    ''' <param name="p_strSysDate"></param>
    ''' <param name="p_strStaffDataID"></param>
    ''' <param name="p_intCurrentDate"></param>
    ''' <returns></returns>
    Public Function update_NS_KINMUPLAN_F_01(ByVal p_strKinmuCD As String,
                                             ByVal p_strRiyuKBN As String,
                                             ByVal p_strComment As String,
                                             ByVal p_strSysDate As String,
                                             ByVal p_strStaffDataID As String,
                                             ByVal p_intCurrentDate As Integer) As Boolean
        update_NS_KINMUPLAN_F_01 = False

        Try

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "Update NS_KINMUPLAN_F Set"
                w_Sql = w_Sql & " KinmuCD = '" & p_strKinmuCD & "'"
                w_Sql = w_Sql & ", ReasonKbn = '" & p_strRiyuKBN & "'"
                w_Sql = w_Sql & ", HOPECOMMENT = '" & p_strComment & "' " '2015/04/10 Bando Add
                w_Sql = w_Sql & ", LastUpdTimeDate = " & p_strSysDate
                w_Sql = w_Sql & ", RegistrantID = '" & General.g_strUserID & "'"
                w_Sql = w_Sql & " Where StaffMngID = '" & p_strStaffDataID & "'"
                w_Sql = w_Sql & " And DateF = " & Convert.ToString(p_intCurrentDate)
                w_Sql = w_Sql & " And HospitalCD = '" & General.g_strHospitalCD & "'"
                w_Sql = w_Sql & " "
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "Update NS_KINMUPLAN_F Set"
                w_Sql = w_Sql & " KinmuCD = '" & p_strKinmuCD & "'"
                w_Sql = w_Sql & ", ReasonKbn = '" & p_strRiyuKBN & "'"
                w_Sql = w_Sql & ", HOPECOMMENT = '" & p_strComment & "' " '2015/04/10 Bando Add
                w_Sql = w_Sql & ", LastUpdTimeDate = " & p_strSysDate
                w_Sql = w_Sql & ", RegistrantID = '" & General.g_strUserID & "'"
                w_Sql = w_Sql & " Where StaffMngID = '" & p_strStaffDataID & "'"
                w_Sql = w_Sql & " And DateF = " & Convert.ToString(p_intCurrentDate)
                w_Sql = w_Sql & " And HospitalCD = '" & General.g_strHospitalCD & "'"
                w_Sql = w_Sql & " "
            End If

            Call General.paDBExecute(w_Sql)
            update_NS_KINMUPLAN_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="p_strKinmuCD"></param>
    ''' <param name="p_strRiyuKBN"></param>
    ''' <param name="p_strKangoCD"></param>
    ''' <param name="p_strComment"></param>
    ''' <param name="p_strSysDate"></param>
    ''' <param name="p_intCurrentDate"></param>
    ''' <param name="p_strStaffID"></param>
    ''' <returns></returns>
    Public Function update_NS_KINMUPLAN_F_02(ByVal p_strKinmuCD As String,
                                             ByVal p_strRiyuKBN As String,
                                             ByVal p_strKangoCD As String,
                                             ByVal p_strComment As String,
                                             ByVal p_strSysDate As String,
                                             ByVal p_intCurrentDate As Integer,
                                             ByVal p_strStaffID As String) As Boolean
        update_NS_KINMUPLAN_F_02 = False

        Try

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "UPDATE NS_KINMUPLAN_F SET "
                w_Sql = w_Sql & " KINMUCD = '" & p_strKinmuCD & "' "
                w_Sql = w_Sql & ",REASONKBN = '" & p_strRiyuKBN & "' "
                w_Sql = w_Sql & ",OUENKINMUDEPTCD = '" & p_strKangoCD & "' "
                w_Sql = w_Sql & ",HOPECOMMENT = '" & p_strComment & "' " '2015/04/10 Bando Add
                w_Sql = w_Sql & ",LASTUPDTIMEDATE = " & p_strSysDate & " "
                w_Sql = w_Sql & ",REGISTRANTID = '" & General.g_strUserID & "' "
                w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "AND DATEF = " & p_intCurrentDate & " "
                w_Sql = w_Sql & "AND STAFFMNGID = '" & p_strStaffID & "' "
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "UPDATE NS_KINMUPLAN_F SET "
                w_Sql = w_Sql & " KINMUCD = '" & p_strKinmuCD & "' "
                w_Sql = w_Sql & ",REASONKBN = '" & p_strRiyuKBN & "' "
                w_Sql = w_Sql & ",OUENKINMUDEPTCD = '" & p_strKangoCD & "' "
                w_Sql = w_Sql & ",HOPECOMMENT = '" & p_strComment & "' " '2015/04/10 Bando Add
                w_Sql = w_Sql & ",LASTUPDTIMEDATE = " & p_strSysDate & " "
                w_Sql = w_Sql & ",REGISTRANTID = '" & General.g_strUserID & "' "
                w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "AND DATEF = " & p_intCurrentDate & " "
                w_Sql = w_Sql & "AND STAFFMNGID = '" & p_strStaffID & "' "
            End If

            Call General.paDBExecute(w_Sql)
            update_NS_KINMUPLAN_F_02 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="p_strKinmuCD"></param>
    ''' <param name="p_strRiyuKBN"></param>
    ''' <param name="p_strComment"></param>
    ''' <param name="p_strSysDate"></param>
    ''' <param name="p_intCurrentDate"></param>
    ''' <param name="p_strStaffDataID"></param>
    ''' <returns></returns>
    Public Function update_NS_KINMUPLAN_F_03(ByVal p_strKinmuCD As String,
                                             ByVal p_strRiyuKBN As String,
                                             ByVal p_strComment As String,
                                             ByVal p_strSysDate As String,
                                             ByVal p_intCurrentDate As Integer,
                                             ByVal p_strStaffDataID As String) As Boolean
        update_NS_KINMUPLAN_F_03 = False
        Try
            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "Update NS_KINMUPLAN_F SET"
                w_Sql = w_Sql & " KINMUCD = '" & p_strKinmuCD & "'"
                w_Sql = w_Sql & ", REASONKBN = '" & p_strRiyuKBN & "'"
                w_Sql = w_Sql & ", HOPECOMMENT = '" & p_strComment & "'" '2015/04/10 Bando Add
                w_Sql = w_Sql & ", LASTUPDTIMEDATE = " & p_strSysDate
                w_Sql = w_Sql & ", REGISTRANTID = '" & General.g_strUserID & "'"
                w_Sql = w_Sql & " WHERE HOSPITALCD = '" & General.g_strHospitalCD & "'"
                w_Sql = w_Sql & " AND DATEF = " & Convert.ToString(p_intCurrentDate)
                w_Sql = w_Sql & " AND STAFFMNGID = '" & p_strStaffDataID & "'"
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "Update NS_KINMUPLAN_F SET"
                w_Sql = w_Sql & " KINMUCD = '" & p_strKinmuCD & "'"
                w_Sql = w_Sql & ", REASONKBN = '" & p_strRiyuKBN & "'"
                w_Sql = w_Sql & ", HOPECOMMENT = '" & p_strComment & "'" '2015/04/10 Bando Add
                w_Sql = w_Sql & ", LASTUPDTIMEDATE = " & p_strSysDate
                w_Sql = w_Sql & ", REGISTRANTID = '" & General.g_strUserID & "'"
                w_Sql = w_Sql & " WHERE HOSPITALCD = '" & General.g_strHospitalCD & "'"
                w_Sql = w_Sql & " AND DATEF = " & Convert.ToString(p_intCurrentDate)
                w_Sql = w_Sql & " AND STAFFMNGID = '" & p_strStaffDataID & "'"
            End If

            Call General.paDBExecute(w_Sql)
            update_NS_KINMUPLAN_F_03 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="p_strKinmuCD"></param>
    ''' <param name="p_strRiyuKBN"></param>
    ''' <param name="p_strKangoCD"></param>
    ''' <param name="p_strComment"></param>
    ''' <param name="p_strSysDate"></param>
    ''' <param name="p_intCurrentDate"></param>
    ''' <param name="p_strStaffID"></param>
    ''' <returns></returns>
    Public Function update_NS_KINMURESULT_F_01(ByVal p_strKinmuCD As String,
                                               ByVal p_strRiyuKBN As String,
                                               ByVal p_strKangoCD As String,
                                               ByVal p_strComment As String,
                                               ByVal p_strSysDate As String,
                                               ByVal p_intCurrentDate As Integer,
                                               ByVal p_strStaffID As String) As Boolean
        update_NS_KINMURESULT_F_01 = False

        Try
            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "Update NS_KINMURESULT_F set "
                w_Sql = w_Sql & " KinmuCD = '" & p_strKinmuCD & "' "
                w_Sql = w_Sql & ",ReasonKbn = '" & p_strRiyuKBN & "' "
                w_Sql = w_Sql & ",OUENKINMUDEPTCD = '" & p_strKangoCD & "' "
                w_Sql = w_Sql & ",HOPECOMMENT = '" & p_strComment & "' " '2015/04/10 Bando Add
                w_Sql = w_Sql & ",LastUpdTimeDate = " & p_strSysDate & " "
                w_Sql = w_Sql & ",RegistrantID = '" & General.g_strUserID & "' "
                w_Sql = w_Sql & "Where HospitalCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "And DateF = " & p_intCurrentDate & " "
                w_Sql = w_Sql & "And StaffMngID = '" & p_strStaffID & "' "
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "Update NS_KINMURESULT_F set "
                w_Sql = w_Sql & " KinmuCD = '" & p_strKinmuCD & "' "
                w_Sql = w_Sql & ",ReasonKbn = '" & p_strRiyuKBN & "' "
                w_Sql = w_Sql & ",OUENKINMUDEPTCD = '" & p_strKangoCD & "' "
                w_Sql = w_Sql & ",HOPECOMMENT = '" & p_strComment & "' " '2015/04/10 Bando Add
                w_Sql = w_Sql & ",LastUpdTimeDate = " & p_strSysDate & " "
                w_Sql = w_Sql & ",RegistrantID = '" & General.g_strUserID & "' "
                w_Sql = w_Sql & "Where HospitalCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "And DateF = " & p_intCurrentDate & " "
                w_Sql = w_Sql & "And StaffMngID = '" & p_strStaffID & "' "
            End If
            Call General.paDBExecute(w_Sql)
            update_NS_KINMURESULT_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="p_strKinmuCD"></param>
    ''' <param name="p_strRiyuKBN"></param>
    ''' <param name="p_strKangoCD"></param>
    ''' <param name="p_blnSaveType"></param>
    ''' <param name="p_strComment"></param>
    ''' <param name="p_strSysDate"></param>
    ''' <param name="p_strStaffID"></param>
    ''' <param name="p_strCurrentDate"></param>
    ''' <returns></returns>
    Public Function update_NS_KINMURESULT_F_03(ByVal p_strKinmuCD As String,
                                               ByVal p_strRiyuKBN As String,
                                               ByVal p_strKangoCD As String,
                                               ByVal p_blnSaveType As Boolean,
                                               ByVal p_strComment As String,
                                               ByVal p_strSysDate As String,
                                               ByVal p_strStaffID As String,
                                               ByVal p_strCurrentDate As String) As Boolean
        update_NS_KINMURESULT_F_03 = False

        Try

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "Update NS_KINMURESULT_F Set"
                w_Sql = w_Sql & " KinmuCD = '" & p_strKinmuCD & "'"
                w_Sql = w_Sql & ", ReasonKbn = '" & p_strRiyuKBN & "'"
                w_Sql = w_Sql & ", OUENKINMUDEPTCD = '" & p_strKangoCD & "'"
                If p_blnSaveType = True Then
                    w_Sql = w_Sql & ", DecDeptCD = '" & General.g_strSelKinmuDeptCD & "'"
                End If
                w_Sql = w_Sql & ", HOPECOMMENT = '" & p_strComment & "'" '2015/04/10 Bando Add
                w_Sql = w_Sql & ", LastUpdTimeDate = " & p_strSysDate
                w_Sql = w_Sql & ", RegistrantID = '" & General.g_strUserID & "'"
                w_Sql = w_Sql & " Where StaffMngID = '" & p_strStaffID & "'"
                w_Sql = w_Sql & " And DateF = " & p_strCurrentDate
                w_Sql = w_Sql & " And HospitalCD = '" & General.g_strHospitalCD & "'"
                w_Sql = w_Sql & " "
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "Update NS_KINMURESULT_F Set"
                w_Sql = w_Sql & " KinmuCD = '" & p_strKinmuCD & "'"
                w_Sql = w_Sql & ", ReasonKbn = '" & p_strRiyuKBN & "'"
                w_Sql = w_Sql & ", OUENKINMUDEPTCD = '" & p_strKangoCD & "'"
                If p_blnSaveType = True Then
                    w_Sql = w_Sql & ", DecDeptCD = '" & General.g_strSelKinmuDeptCD & "'"
                End If
                w_Sql = w_Sql & ", HOPECOMMENT = '" & p_strComment & "'" '2015/04/10 Bando Add
                w_Sql = w_Sql & ", LastUpdTimeDate = " & p_strSysDate
                w_Sql = w_Sql & ", RegistrantID = '" & General.g_strUserID & "'"
                w_Sql = w_Sql & " Where StaffMngID = '" & p_strStaffID & "'"
                w_Sql = w_Sql & " And DateF = " & p_strCurrentDate
                w_Sql = w_Sql & " And HospitalCD = '" & General.g_strHospitalCD & "'"
                w_Sql = w_Sql & " "
            End If

            Call General.paDBExecute(w_Sql)
            update_NS_KINMURESULT_F_03 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="p_intSysDate"></param>
    ''' <param name="p_strPlanNo"></param>
    ''' <returns></returns>
    Public Function update_NS_PLANDECISION_F_01(ByVal p_intSysDate As String,
                                                ByVal p_strPlanNo As String) As Boolean
        update_NS_PLANDECISION_F_01 = False

        Try

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "UPDATE NS_PLANDECISION_F "
                w_Sql = w_Sql & "SET LASTUPDTIMEDATE = " & p_intSysDate & ", "
                w_Sql = w_Sql & "    REGISTRANTID = '" & General.g_strUserID & "' "
                w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "AND   KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "' "
                w_Sql = w_Sql & "AND   PLANNO = " & p_strPlanNo & " "
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "UPDATE NS_PLANDECISION_F "
                w_Sql = w_Sql & "SET LASTUPDTIMEDATE = " & p_intSysDate & ", "
                w_Sql = w_Sql & "    REGISTRANTID = '" & General.g_strUserID & "' "
                w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "AND   KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "' "
                w_Sql = w_Sql & "AND   PLANNO = " & p_strPlanNo & " "
            End If
            Call General.paDBExecute(w_Sql)
            update_NS_PLANDECISION_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="p_objStaffData"></param>
    ''' <param name="p_intIndex"></param>
    ''' <param name="p_intDispNo1"></param>
    ''' <param name="p_intDispNo2"></param>
    ''' <param name="p_intDispNo3"></param>
    ''' <param name="p_intDispNo4"></param>
    ''' <param name="p_intDispNo5"></param>
    ''' <param name="p_dblSysDate"></param>
    ''' <param name="p_intPlanNo"></param>
    ''' <returns></returns>
    Public Function update_NS_STAFFHISTORY_F_01(ByVal p_objStaffData As Object,
                                                ByVal p_intIndex As Integer,
                                                ByVal p_intDispNo1 As Integer,
                                                ByVal p_intDispNo2 As Integer,
                                                ByVal p_intDispNo3 As Integer,
                                                ByVal p_intDispNo4 As Integer,
                                                ByVal p_intDispNo5 As Integer,
                                                ByVal p_dblSysDate As Double,
                                                ByVal p_intPlanNo As Integer) As Boolean
        update_NS_STAFFHISTORY_F_01 = False

        Try
            w_sqlBuilder = New System.Text.StringBuilder

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "UPDATE NS_STAFFHISTORY_F SET "
                w_Sql = w_Sql & " SKILLLVLCD = '" & p_objStaffData(p_intIndex).GiryoLvCD & "' "
                w_Sql = w_Sql & ",DISPNO = '" & p_objStaffData(p_intIndex).HyojiNo & "' "
                w_Sql = w_Sql & ",DISPNO1 = '" & p_intDispNo1 & "' "
                w_Sql = w_Sql & ",DISPNO2 = '" & p_intDispNo2 & "' "
                w_Sql = w_Sql & ",DISPNO3 = '" & p_intDispNo3 & "' "
                w_Sql = w_Sql & ",DISPNO4 = '" & p_intDispNo4 & "' "
                w_Sql = w_Sql & ",DISPNO5 = '" & p_intDispNo5 & "' "
                w_Sql = w_Sql & ",AUTOALLOCKBN = '" & p_objStaffData(p_intIndex).AutoKBN & "' "
                w_Sql = w_Sql & ",TEAM = " & p_objStaffData(p_intIndex).Team & " "
                w_Sql = w_Sql & ",NIGHTONLYSTAFFKBN = " & p_objStaffData(p_intIndex).YakinKBN & " "
                w_Sql = w_Sql & ",PATTERNCD = '" & p_objStaffData(p_intIndex).PatternCD & "' "
                w_Sql = w_Sql & ",LASTUPDTIMEDATE = " & p_dblSysDate & " "
                w_Sql = w_Sql & ",REGISTRANTID = '" & General.g_strUserID & "' "
                w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "AND PLANNO = " & p_intPlanNo & " "
                w_Sql = w_Sql & "AND KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "' "
                w_Sql = w_Sql & "AND STAFFMNGID = '" & p_objStaffData(p_intIndex).ID & "' "
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "UPDATE NS_STAFFHISTORY_F SET "
                w_Sql = w_Sql & " SKILLLVLCD = '" & p_objStaffData(p_intIndex).GiryoLvCD & "' "
                w_Sql = w_Sql & ",DISPNO = '" & p_objStaffData(p_intIndex).HyojiNo & "' "
                w_Sql = w_Sql & ",DISPNO1 = '" & p_intDispNo1 & "' "
                w_Sql = w_Sql & ",DISPNO2 = '" & p_intDispNo2 & "' "
                w_Sql = w_Sql & ",DISPNO3 = '" & p_intDispNo3 & "' "
                w_Sql = w_Sql & ",DISPNO4 = '" & p_intDispNo4 & "' "
                w_Sql = w_Sql & ",DISPNO5 = '" & p_intDispNo5 & "' "
                w_Sql = w_Sql & ",AUTOALLOCKBN = '" & p_objStaffData(p_intIndex).AutoKBN & "' "
                w_Sql = w_Sql & ",TEAM = " & p_objStaffData(p_intIndex).Team & " "
                w_Sql = w_Sql & ",NIGHTONLYSTAFFKBN = " & p_objStaffData(p_intIndex).YakinKBN & " "
                w_Sql = w_Sql & ",PATTERNCD = '" & p_objStaffData(p_intIndex).PatternCD & "' "
                w_Sql = w_Sql & ",LASTUPDTIMEDATE = " & p_dblSysDate & " "
                w_Sql = w_Sql & ",REGISTRANTID = '" & General.g_strUserID & "' "
                w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "AND PLANNO = " & p_intPlanNo & " "
                w_Sql = w_Sql & "AND KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "' "
                w_Sql = w_Sql & "AND STAFFMNGID = '" & p_objStaffData(p_intIndex).ID & "' "
            End If

            Call General.paDBExecute(w_Sql)
            update_NS_STAFFHISTORY_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="p_strStaffDataID"></param>
    ''' <param name="p_intCurrentDate"></param>
    ''' <returns></returns>
    Public Function delete_NS_KINMUPLAN_F_01(ByVal p_strStaffDataID As String,
                                             ByVal p_intCurrentDate As Integer) As Boolean
        delete_NS_KINMUPLAN_F_01 = False

        Try

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "Delete From NS_KINMUPLAN_F"
                w_Sql = w_Sql & " Where StaffMngID = '" & p_strStaffDataID & "'"
                w_Sql = w_Sql & " And DateF = " & Convert.ToString(p_intCurrentDate)
                w_Sql = w_Sql & " And HospitalCD = '" & General.g_strHospitalCD & "'"
                w_Sql = w_Sql & " "
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "Delete From NS_KINMUPLAN_F"
                w_Sql = w_Sql & " Where StaffMngID = '" & p_strStaffDataID & "'"
                w_Sql = w_Sql & " And DateF = " & Convert.ToString(p_intCurrentDate)
                w_Sql = w_Sql & " And HospitalCD = '" & General.g_strHospitalCD & "'"
                w_Sql = w_Sql & " "
            End If

            Call General.paDBExecute(w_Sql)
            delete_NS_KINMUPLAN_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="p_intDate"></param>
    ''' <param name="p_strStaffID"></param>
    ''' <returns></returns>
    Public Function delete_NS_NENKYU_F_01(ByVal p_intDate As Integer,
                                          ByVal p_strStaffID As String) As Boolean
        delete_NS_NENKYU_F_01 = False
        Try
            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "Delete From NS_NENKYU_F "
                w_Sql = w_Sql & "Where HospitalCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "And DateF = " & Convert.ToString(p_intDate) & " "
                w_Sql = w_Sql & "And StaffMngID = '" & p_strStaffID & "' "
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "Delete From NS_NENKYU_F "
                w_Sql = w_Sql & "Where HospitalCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "And DateF = " & Convert.ToString(p_intDate) & " "
                w_Sql = w_Sql & "And StaffMngID = '" & p_strStaffID & "' "
            End If

            Call General.paDBExecute(w_Sql)
            delete_NS_NENKYU_F_01 = True

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="p_intDateF"></param>
    ''' <param name="p_strStaffMngID"></param>
    ''' <returns></returns>
    Public Function delete_NS_KINMURESULT_F_01(ByVal p_intDateF As Integer,
                                               ByVal p_strStaffMngID As String) As Boolean
        delete_NS_KINMURESULT_F_01 = False

        Try
            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "DELETE FROM NS_KINMURESULT_F "
                w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "AND DATEF = " & p_intDateF & " "
                w_Sql = w_Sql & "AND STAFFMNGID = '" & p_strStaffMngID & "' "
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "DELETE FROM NS_KINMURESULT_F "
                w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "AND DATEF = " & p_intDateF & " "
                w_Sql = w_Sql & "AND STAFFMNGID = '" & p_strStaffMngID & "' "
            End If

            Call General.paDBExecute(w_Sql)
            delete_NS_KINMURESULT_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="p_intCurrentDate"></param>
    ''' <param name="p_strStaffID"></param>
    ''' <returns></returns>
    Public Function delete_NS_KINMUCHGHISTORY_F_01(ByVal p_intCurrentDate As Integer,
                                                   ByVal p_strStaffID As String) As Boolean
        delete_NS_KINMUCHGHISTORY_F_01 = False

        Try
            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "DELETE FROM NS_KINMUCHGHISTORY_F "
                w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "AND DATEF = " & p_intCurrentDate & " "
                w_Sql = w_Sql & "AND STAFFMNGID = '" & p_strStaffID & "' "
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "DELETE FROM NS_KINMUCHGHISTORY_F "
                w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "AND DATEF = " & p_intCurrentDate & " "
                w_Sql = w_Sql & "AND STAFFMNGID = '" & p_strStaffID & "' "
            End If

            Call General.paDBExecute(w_Sql)
            delete_NS_KINMUCHGHISTORY_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="p_intPlanNo"></param>
    ''' <returns></returns>
    Public Function delete_NS_PLANDECISION_F_01(ByVal p_intPlanNo As Integer) As Boolean
        delete_NS_PLANDECISION_F_01 = False

        Try

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "DELETE FROM NS_PLANDECISION_F "
                w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "AND   KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "' "
                w_Sql = w_Sql & "AND   PLANNO = " & p_intPlanNo & " "
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "DELETE FROM NS_PLANDECISION_F "
                w_Sql = w_Sql & "WHERE HOSPITALCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "AND   KINMUDEPTCD = '" & General.g_strSelKinmuDeptCD & "' "
                w_Sql = w_Sql & "AND   PLANNO = " & p_intPlanNo & " "
            End If
            Call General.paDBExecute(w_Sql)
            delete_NS_PLANDECISION_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="p_objStaffData"></param>
    ''' <param name="p_intIndex"></param>
    ''' <param name="p_intDaikyuStartDate"></param>
    ''' <param name="p_intDaikyuEndDate"></param>
    ''' <returns></returns>
    Public Function delete_NS_DAIKYUMNG_F_01(ByVal p_objStaffData As Object,
                                             ByVal p_intIndex As Integer,
                                             ByVal p_intDaikyuStartDate As Integer,
                                             ByVal p_intDaikyuEndDate As Integer) As Boolean
        delete_NS_DAIKYUMNG_F_01 = False

        Try
            w_sqlBuilder = New System.Text.StringBuilder

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "Delete From NS_DAIKYUMNG_F "
                w_Sql = w_Sql & "Where HospitalCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "And StaffMngID = '" & p_objStaffData(p_intIndex).ID & "' "
                w_Sql = w_Sql & "And WorkHolKinmuDate >= " & p_intDaikyuStartDate & " "
                w_Sql = w_Sql & "And WorkHolKinmuDate <= " & p_intDaikyuEndDate
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "Delete From NS_DAIKYUMNG_F "
                w_Sql = w_Sql & "Where HospitalCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "And StaffMngID = '" & p_objStaffData(p_intIndex).ID & "' "
                w_Sql = w_Sql & "And WorkHolKinmuDate >= " & p_intDaikyuStartDate & " "
                w_Sql = w_Sql & "And WorkHolKinmuDate <= " & p_intDaikyuEndDate
            End If

            Call General.paDBExecute(w_Sql)
            delete_NS_DAIKYUMNG_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HA
    ''' </summary>
    ''' <param name="p_objStaffData"></param>
    ''' <param name="p_intIndex"></param>
    ''' <param name="p_intDaikyuStartDate"></param>
    ''' <param name="p_intDaikyuEndDate"></param>
    ''' <returns></returns>
    Public Function delete_NS_DAIKYUDETAILMNG_F_01(ByVal p_objStaffData As Object,
                                                   ByVal p_intIndex As Integer,
                                                   ByVal p_intDaikyuStartDate As Integer,
                                                   ByVal p_intDaikyuEndDate As Integer) As Boolean
        delete_NS_DAIKYUDETAILMNG_F_01 = False

        Try

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then      'ORACLE
                w_Sql = "Delete From NS_DAIKYUDETAILMNG_F "
                w_Sql = w_Sql & "Where HospitalCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "And StaffMngID = '" & p_objStaffData(p_intIndex).ID & "' "
                w_Sql = w_Sql & "And WorkHolKinmuDate >= " & p_intDaikyuStartDate & " "
                w_Sql = w_Sql & "And WorkHolKinmuDate <= " & p_intDaikyuEndDate
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then          'SQL
                w_Sql = "Delete From NS_DAIKYUDETAILMNG_F "
                w_Sql = w_Sql & "Where HospitalCD = '" & General.g_strHospitalCD & "' "
                w_Sql = w_Sql & "And StaffMngID = '" & p_objStaffData(p_intIndex).ID & "' "
                w_Sql = w_Sql & "And WorkHolKinmuDate >= " & p_intDaikyuStartDate & " "
                w_Sql = w_Sql & "And WorkHolKinmuDate <= " & p_intDaikyuEndDate
            End If

            Call General.paDBExecute(w_Sql)
            delete_NS_DAIKYUDETAILMNG_F_01 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' NSK0000HJ
    ''' </summary>
    ''' <param name="p_intSelDate"></param>
    ''' <param name="p_strMngStaffID"></param>
    ''' <returns></returns>
    Public Function delete_NS_DAIKYUDETAILMNG_F_02(ByVal p_intSelDate As Integer,
                                                   ByVal p_strMngStaffID As String) As Boolean
        Try
            delete_NS_DAIKYUDETAILMNG_F_02 = False
            w_sqlBuilder = New System.Text.StringBuilder

            If General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_PassThrough Then 'ORACLE
                w_sqlBuilder.Append("Delete From NS_DAIKYUDETAILMNG_F ")
                w_sqlBuilder.Append(" where WorkHolKinmuDate = ")
                w_sqlBuilder.Append(p_intSelDate)
                w_sqlBuilder.Append(" and StaffMngID = '")
                w_sqlBuilder.Append(p_strMngStaffID)
                w_sqlBuilder.Append("' and HospitalCD = '")
                w_sqlBuilder.Append(General.g_strHospitalCD)
                w_sqlBuilder.Append("'")
            ElseIf General.BasGeneral.g_InstallType = General.BasGeneral.gInstall_Enum.AccessType_SQL Then 'SQL Server
                w_sqlBuilder.Append("Delete From NS_DAIKYUDETAILMNG_F ")
                w_sqlBuilder.Append(" where WorkHolKinmuDate = ")
                w_sqlBuilder.Append(p_intSelDate)
                w_sqlBuilder.Append(" and StaffMngID = '")
                w_sqlBuilder.Append(p_strMngStaffID)
                w_sqlBuilder.Append("' and HospitalCD = '")
                w_sqlBuilder.Append(General.g_strHospitalCD)
                w_sqlBuilder.Append("'")
            End If

            w_Sql = w_sqlBuilder.ToString
            Call General.paDBExecute(w_Sql)
            delete_NS_DAIKYUDETAILMNG_F_02 = True
        Catch ex As Exception
            Throw
        End Try
    End Function

End Class
