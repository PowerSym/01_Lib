
! This Model Source Code is for UserCode - <27/04/2020>
C2C**************************************************************************************
C2C MODIFICATION HISTORY:
C2C     27/APR/2020: DIDNG first created
C2C--------------------------------------------------------------------------------------
      SUBROUTINE USRCD1(KM,ISLOT)
C2C
      INCLUDE 'COMON4.INS'
C2C
      IMPLICIT NONE
      INTRINSIC Abs, Max, Min
C++
C2C   Declare variables in PSSE 
C--
      INTEGER    KM  
      INTEGER    ISLOT                 
      INTEGER    iUSRCD1_Strt_Con  
      INTEGER    iUSRCD1_Strt_State
      INTEGER    iUSRCD1_Strt_Var  
      INTEGER    iUSRCD1_Strt_Icon
	  
      INTEGER    iNum_Con
      INTEGER    iNum_State
      INTEGER    iNum_Var
      INTEGER    iNum_Icon
	  
      INTEGER    iConDecs
      INTEGER    iIconDecs
      
	  INTEGER    iBusNbr                 
      INTEGER    iPPC_bus
      Character  strMachID
	  
	  INTEGER    iMSU1_Fallback_Enable 
	  INTEGER    iFreq_Injection_Enable
	  INTEGER    iMSU1_Discharge_VAR
	  INTEGER    iMSU2_Discharge_VAR
	  INTEGER    iMSU3_Discharge_VAR
	  INTEGER    iMSU4_Discharge_VAR
 
	  
      REAL       rP1111_pu
      REAL       rP0111_pu
      REAL       rP1011_pu
      REAL       rP1101_pu
      REAL       rP1110_pu
      REAL       rP0011_pu
      REAL       rP1001_pu
      REAL       rP1100_pu
      REAL       rP1000_pu
      REAL       rP0100_pu
      REAL       rP0010_pu
      REAL       rP0001_pu
      REAL       rP0000_pu
      REAL       rFreq_Str_Time_s         
      REAL       rFreq_End_Time_s         
      REAL       rROCOF1_pus       
      REAL       rROCOF2_pus       
      REAL       rFreq_Change_Up_Limit_pu 
      REAL       rFreq_Change_Low_Limit_pu
	  
      INTEGER    iPPC_P_Ctrl_CON      
      INTEGER    iPPC_Q_Ctrl_VAR
      INTEGER    iPPC_Q_Ctrl_CON
      INTEGER    iPPC_P_Ctrl_STATE
      
	  REAL       rMSU_Discharging_Time_s
      REAL       rCurrent_MSU1_Discharging_Time_s
      REAL       rCurrent_MSU2_Discharging_Time_s
      REAL       rCurrent_MSU3_Discharging_Time_s
      REAL       rCurrent_MSU4_Discharging_Time_s

      REAL       r_MSU1_Discharging_Time_Temp_s
      REAL       r_MSU2_Discharging_Time_Temp_s
      REAL       r_MSU3_Discharging_Time_Temp_s
      REAL       r_MSU4_Discharging_Time_Temp_s

      REAL       r_MSU1_availability
      REAL       r_MSU2_availability
      REAL       r_MSU3_availability
      REAL       r_MSU4_availability	

      REAL       rPwf_limit_pu
      REAL       rTemp
      REAL       rCurrent_Freq_Dev_pu

 	  
      INTEGER    IERR
      INTEGER    iIndex
	  
C2C
      Integer, Parameter:: Array_Character_Number = 50
	  
	  CHARACTER*30,   chCON_ARRAY(Array_Character_Number)
C2C
C++
C********************* START MAIN PSSE PROGRAM: MODE 8 ********************************
C--      
      IF (MODE == 8) THEN                        
C++
C2C   Add CON_DSCRPT & ICON_DSCRPT
       iConDecs                  = 1
       CON_DSCRPT(iConDecs+00)   = 'PLimit_1111, pu'
       CON_DSCRPT(iConDecs+01)   = 'PLimit_0111, pu'
       CON_DSCRPT(iConDecs+02)   = 'PLimit_1011, pu'
       CON_DSCRPT(iConDecs+03)   = 'PLimit_1101, pu'
       CON_DSCRPT(iConDecs+04)   = 'PLimit_1110, pu'
       CON_DSCRPT(iConDecs+05)   = 'PLimit_0011, pu'
       CON_DSCRPT(iConDecs+06)   = 'PLimit_1001, pu'
       CON_DSCRPT(iConDecs+07)   = 'PLimit_1100, pu'
       CON_DSCRPT(iConDecs+08)   = 'PLimit_1000, pu'
       CON_DSCRPT(iConDecs+09)   = 'PLimit_0100, pu'
       CON_DSCRPT(iConDecs+10)   = 'PLimit_0010, pu'
       CON_DSCRPT(iConDecs+11)   = 'PLimit_0001, pu'
       CON_DSCRPT(iConDecs+12)   = 'PLimit_0000, pu'
       CON_DSCRPT(iConDecs+13)   = 'Reserved'
       CON_DSCRPT(iConDecs+14)   = 'Reserved'
       CON_DSCRPT(iConDecs+15)   = 'Freq inject start time, s'
       CON_DSCRPT(iConDecs+16)   = 'Freq inject end time, s'
       CON_DSCRPT(iConDecs+17)   = 'ROCOF 1, pu/s'
       CON_DSCRPT(iConDecs+18)   = 'ROCOF 2, pu/s'
       CON_DSCRPT(iConDecs+19)   = 'Freq deviation inject up limit, pu'
	   CON_DSCRPT(iConDecs+20)   = 'Freq deviation inject low limit, pu'
	   CON_DSCRPT(iConDecs+21)   = 'Reserved'
	   CON_DSCRPT(iConDecs+22)   = 'Reserved'
	   CON_DSCRPT(iConDecs+23)   = 'Reserved'
	   CON_DSCRPT(iConDecs+24)   = 'Reserved'
	   CON_DSCRPT(iConDecs+25)   = 'Reserved'
	   CON_DSCRPT(iConDecs+26)   = 'Reserved'
	   CON_DSCRPT(iConDecs+27)   = 'Reserved'
	   CON_DSCRPT(iConDecs+28)   = 'Reserved'
	   CON_DSCRPT(iConDecs+29)   = 'Reserved'

C2C
       iIconDecs                 = 1
       ICON_DSCRPT(iIconDecs+00) = 'PPC attached bus'
       ICON_DSCRPT(iIconDecs+01) = 'PPC Machine ID'
	   ICON_DSCRPT(iIconDecs+02) = 'Enable/Disable MSU fallback control'
       ICON_DSCRPT(iIconDecs+03) = 'Enable/Disable frequency injection'
	   ICON_DSCRPT(iIconDecs+04) = 'Reserved' 
       ICON_DSCRPT(iIconDecs+05) = 'Reserved'
	   ICON_DSCRPT(iIconDecs+06) = 'MSU 1 Discharging Time VAR #' 
       ICON_DSCRPT(iIconDecs+07) = 'MSU 2 Discharging Time VAR #'
	   ICON_DSCRPT(iIconDecs+08) = 'MSU 3 Discharging Time VAR #'
	   ICON_DSCRPT(iIconDecs+09) = 'MSU 4 Discharging Time VAR #'
	   ICON_DSCRPT(iIconDecs+10) = 'Reserved' 
       ICON_DSCRPT(iIconDecs+11) = 'Reserved'
       ICON_DSCRPT(iIconDecs+12) = 'Reserved'
       ICON_DSCRPT(iIconDecs+13) = 'Reserved'
       ICON_DSCRPT(iIconDecs+14) = 'Reserved'
       ICON_DSCRPT(iIconDecs+15) = 'Reserved'
       ICON_DSCRPT(iIconDecs+16) = 'Reserved'
       ICON_DSCRPT(iIconDecs+17) = 'Reserved'
       ICON_DSCRPT(iIconDecs+18) = 'Reserved'
       ICON_DSCRPT(iIconDecs+19) = 'Reserved'
C--

       RETURN
      ENDIF
C**************************************************************************************           
C2C
      iBusNbr            = NUMBUS(KM)
      iUSRCD1_Strt_Con   = STRTCCT(1, ISLOT)
      iUSRCD1_Strt_State = STRTCCT(2, ISLOT)
      iUSRCD1_Strt_Var   = STRTCCT(3, ISLOT)
      iUSRCD1_Strt_Icon  = STRTCCT(4, ISLOT)
	  
      iPPC_bus           = ICON(iUSRCD1_Strt_Icon+0)
	  strMachID          = ChrIcn(iUSRCD1_Strt_Icon+1)
C2C
	  CALL CCTMIND_BUSO(iBusNbr, 'USRCD1', 'NCON', iNum_Con, iErr)
	  CALL CCTMIND_BUSO(iBusNbr, 'USRCD1', 'NSTATE', iNum_State, iErr)
	  CALL CCTMIND_BUSO(iBusNbr, 'USRCD1', 'NVAR', iNum_Var, iErr)
	  CALL CCTMIND_BUSO(iBusNbr, 'USRCD1', 'NICON', iNum_Icon, iErr)
C2C
      iMSU1_Fallback_Enable     = ICON(iUSRCD1_Strt_Icon + 2)
      iFreq_Injection_Enable    = ICON(iUSRCD1_Strt_Icon + 3)
      iMSU1_Discharge_VAR       = ICON(iUSRCD1_Strt_Icon + 6)
      iMSU1_Discharge_VAR       = ICON(iUSRCD1_Strt_Icon + 6)
      iMSU2_Discharge_VAR       = ICON(iUSRCD1_Strt_Icon + 7)
      iMSU3_Discharge_VAR       = ICON(iUSRCD1_Strt_Icon + 8)
      iMSU4_Discharge_VAR       = ICON(iUSRCD1_Strt_Icon + 9)
C2C
      rP1111_pu                 = CON(iUSRCD1_Strt_Con + 0)
      rP0111_pu                 = CON(iUSRCD1_Strt_Con + 1)
      rP1011_pu                 = CON(iUSRCD1_Strt_Con + 2)
      rP1101_pu                 = CON(iUSRCD1_Strt_Con + 3)
      rP1110_pu                 = CON(iUSRCD1_Strt_Con + 4)
      rP0011_pu                 = CON(iUSRCD1_Strt_Con + 5)
      rP1001_pu                 = CON(iUSRCD1_Strt_Con + 6)
      rP1100_pu                 = CON(iUSRCD1_Strt_Con + 7)
      rP1000_pu                 = CON(iUSRCD1_Strt_Con + 8)
      rP0100_pu                 = CON(iUSRCD1_Strt_Con + 9)
      rP0010_pu                 = CON(iUSRCD1_Strt_Con + 10)
      rP0001_pu                 = CON(iUSRCD1_Strt_Con + 11)
      rP0000_pu                 = CON(iUSRCD1_Strt_Con + 12)
	  rFreq_Str_Time_s          = CON(iUSRCD1_Strt_Con + 15) 
	  rFreq_End_Time_s          = CON(iUSRCD1_Strt_Con + 16)
	  rROCOF1_pus               = CON(iUSRCD1_Strt_Con + 17) 
	  rROCOF2_pus               = CON(iUSRCD1_Strt_Con + 18)
	  rFreq_Change_Up_Limit_pu  = CON(iUSRCD1_Strt_Con + 19) 
	  rFreq_Change_Low_Limit_pu = CON(iUSRCD1_Strt_Con + 20)	  
	  
C2C
      Call WindMInd(iPPC_bus,strMachID,'WGUST','CON',iPPC_P_Ctrl_CON,iErr)
      Call WindMInd(iPPC_bus,strMachID,'WAUX','VAR',iPPC_Q_Ctrl_VAR,iErr)
      Call WindMInd(iPPC_bus,strMachID,'WAUX','CON',iPPC_Q_Ctrl_CON,iErr)
      Call WindMInd(iPPC_bus,strMachID,'WGUST','STATE',iPPC_P_Ctrl_STATE,iErr)
	  rMSU_Discharging_Time_s = CON(iPPC_Q_Ctrl_CON + 57)
C2C
C***************************************************************************************
      IF (MODE > 4) GO TO 1000
C++
C************************* MODE 4 - set NINTEG *****************************************
C--
      IF (MODE==4) THEN
!
         IF ((iUSRCD1_Strt_State + iNum_State) > NINTEG) NINTEG = iUSRCD1_Strt_State + iNum_State
!
         RETURN
      ENDIF

C
C++
C****************** MODE 1 - INITIALIZE ************************************************
C--
      IF (MODE.EQ.1) THEN
C2C
	     VAR(iUSRCD1_Strt_Var + 0:iUSRCD1_Strt_Var + iNum_Var) 	 = 0.

         VAR(iUSRCD1_Strt_Var + 0)    = VAR(iPPC_Q_Ctrl_VAR + iMSU1_Discharge_VAR)
         VAR(iUSRCD1_Strt_Var + 1)    = VAR(iPPC_Q_Ctrl_VAR + iMSU2_Discharge_VAR)
         VAR(iUSRCD1_Strt_Var + 2)    = VAR(iPPC_Q_Ctrl_VAR + iMSU3_Discharge_VAR)
         VAR(iUSRCD1_Strt_Var + 3)    = VAR(iPPC_Q_Ctrl_VAR + iMSU4_Discharge_VAR)
		 
		 rTemp                        = 0.
		 STATE(iPPC_P_Ctrl_STATE + 2) = rTemp
		 STORE(iPPC_P_Ctrl_STATE + 2) = rTemp		 
C2C
         RETURN
      END IF
C2C
C++
C******************** MODE 2 - Derivative calculation output ***************************
C--
      IF (MODE==2) THEN
C2C
	
C2C
         RETURN
      ENDIF
C2C
C++
C******************** MODE 3 - SET Model output ****************************************
C--
      IF (MODE==3) THEN
C2C
         rCurrent_MSU1_Discharging_Time_s = VAR(iPPC_Q_Ctrl_VAR + iMSU1_Discharge_VAR)
         rCurrent_MSU2_Discharging_Time_s = VAR(iPPC_Q_Ctrl_VAR + iMSU2_Discharge_VAR)		
         rCurrent_MSU3_Discharging_Time_s = VAR(iPPC_Q_Ctrl_VAR + iMSU3_Discharge_VAR)		
         rCurrent_MSU4_Discharging_Time_s = VAR(iPPC_Q_Ctrl_VAR + iMSU4_Discharge_VAR)
C2C
		 r_MSU1_Discharging_Time_Temp_s   = rCurrent_MSU1_Discharging_Time_s - VAR(iUSRCD1_Strt_Var + 0)
		 r_MSU2_Discharging_Time_Temp_s   = rCurrent_MSU2_Discharging_Time_s - VAR(iUSRCD1_Strt_Var + 1)
		 r_MSU3_Discharging_Time_Temp_s   = rCurrent_MSU3_Discharging_Time_s - VAR(iUSRCD1_Strt_Var + 2)
		 r_MSU4_Discharging_Time_Temp_s   = rCurrent_MSU4_Discharging_Time_s - VAR(iUSRCD1_Strt_Var + 3)
C2C
		 if (abs(r_MSU1_Discharging_Time_Temp_s) > 0.5*DELT) then
			VAR(iUSRCD1_Strt_Var + 10) = 1.0
		 end if
		 if (abs(r_MSU2_Discharging_Time_Temp_s) > 0.5*DELT) then
			VAR(iUSRCD1_Strt_Var + 11) = 1.0
		 end if
		 if (abs(r_MSU3_Discharging_Time_Temp_s) > 0.5*DELT) then
			VAR(iUSRCD1_Strt_Var + 12) = 1.0
		 end if
		 if (abs(r_MSU4_Discharging_Time_Temp_s) > 0.5*DELT) then
			VAR(iUSRCD1_Strt_Var + 13) = 1.0
		 end if		 
C2C
		 if ((TIME - rCurrent_MSU1_Discharging_Time_s < rMSU_Discharging_Time_s).and.
     &       (VAR(iUSRCD1_Strt_Var + 10) == 1.0)) then
			r_MSU1_availability = 0
		 else 
			r_MSU1_availability        = 1
			VAR(iUSRCD1_Strt_Var + 10) = 0
		 end if
		 if ((TIME - rCurrent_MSU2_Discharging_Time_s < rMSU_Discharging_Time_s).and.
     &       (VAR(iUSRCD1_Strt_Var + 11) == 1.0)) then
			r_MSU2_availability = 0
		 else 
			r_MSU2_availability        = 1
			VAR(iUSRCD1_Strt_Var + 11) = 0
		 end if
		 if ((TIME - rCurrent_MSU3_Discharging_Time_s < rMSU_Discharging_Time_s).and.
     &       (VAR(iUSRCD1_Strt_Var + 12) == 1.0)) then
			r_MSU3_availability = 0
		 else 
			r_MSU3_availability        = 1
			VAR(iUSRCD1_Strt_Var + 12) = 0
		 end if
		 if ((TIME - rCurrent_MSU4_Discharging_Time_s < rMSU_Discharging_Time_s).and.
     &       (VAR(iUSRCD1_Strt_Var + 13) == 1.0)) then
			r_MSU4_availability = 0
		 else 
			r_MSU4_availability        = 1
			VAR(iUSRCD1_Strt_Var + 13) = 0
		 end if
C2C
         if ((r_MSU1_availability == 1).and.(r_MSU2_availability == 1).and.
     &       (r_MSU3_availability == 1).and.(r_MSU4_availability == 1)) then
	         rPwf_limit_pu  = rP1111_pu
         end if
		 
         if ((r_MSU1_availability == 0).and.(r_MSU2_availability == 1).and.
     &       (r_MSU3_availability == 1).and.(r_MSU4_availability == 1)) then
	         rPwf_limit_pu  = rP0111_pu
         end if
		 
         if ((r_MSU1_availability == 1).and.(r_MSU2_availability == 0).and.
     &       (r_MSU3_availability == 1).and.(r_MSU4_availability == 1)) then
	         rPwf_limit_pu  = rP1011_pu
         end if

         if ((r_MSU1_availability == 1).and.(r_MSU2_availability == 1).and.
     &       (r_MSU3_availability == 0).and.(r_MSU4_availability == 1)) then
	         rPwf_limit_pu  = rP1101_pu
         end if	

         if ((r_MSU1_availability == 1).and.(r_MSU2_availability == 1).and.
     &       (r_MSU3_availability == 1).and.(r_MSU4_availability == 0)) then
	         rPwf_limit_pu  = rP1110_pu
         end if

         if ((r_MSU1_availability == 0).and.(r_MSU2_availability == 0).and.
     &       (r_MSU3_availability == 1).and.(r_MSU4_availability == 1)) then
	         rPwf_limit_pu  = rP0011_pu
         end if

         if ((r_MSU1_availability == 1).and.(r_MSU2_availability == 0).and.
     &       (r_MSU3_availability == 0).and.(r_MSU4_availability == 1)) then
	         rPwf_limit_pu  = rP1001_pu
         end if

         if ((r_MSU1_availability == 1).and.(r_MSU2_availability == 1).and.
     &       (r_MSU3_availability == 0).and.(r_MSU4_availability == 0)) then
	         rPwf_limit_pu  = rP1100_pu
         end if

         if ((r_MSU1_availability == 1).and.(r_MSU2_availability == 0).and.
     &       (r_MSU3_availability == 0).and.(r_MSU4_availability == 0)) then
	         rPwf_limit_pu  = rP1000_pu
         end if

         if ((r_MSU1_availability == 0).and.(r_MSU2_availability == 1).and.
     &       (r_MSU3_availability == 0).and.(r_MSU4_availability == 0)) then
	         rPwf_limit_pu  = rP0100_pu
         end if

         if ((r_MSU1_availability == 0).and.(r_MSU2_availability == 0).and.
     &       (r_MSU3_availability == 1).and.(r_MSU4_availability == 0)) then
	         rPwf_limit_pu  = rP0010_pu
         end if

         if ((r_MSU1_availability == 0).and.(r_MSU2_availability == 0).and.
     &       (r_MSU3_availability == 0).and.(r_MSU4_availability == 1)) then
	         rPwf_limit_pu  = rP0001_pu
         end if

         if ((r_MSU1_availability == 0).and.(r_MSU2_availability == 0).and.
     &       (r_MSU3_availability == 0).and.(r_MSU4_availability == 0)) then
	         rPwf_limit_pu  = rP0000_pu
         end if
C2C
        rCurrent_Freq_Dev_pu   = VAR(iUSRCD1_Strt_Var + 8)
		if (TIME < rFreq_Str_Time_s) then
			rTemp = 0.
		else if (TIME < rFreq_End_Time_s) then
		    rTemp = Max(Min(rTemp + DELT*rROCOF1_pus, rFreq_Change_Up_Limit_pu), rFreq_Change_Low_Limit_pu)
		else
		    if (abs(rCurrent_Freq_Dev_pu) > 0.001) then
		       rTemp = rTemp + DELT*rROCOF2_pus
			   if ((rCurrent_Freq_Dev_pu < -0.001) .and. (rTemp > 0)) rTemp = 0
			   if ((rCurrent_Freq_Dev_pu > 0.001) .and. (rTemp < 0)) rTemp = 0
			end if
		end if
C2C
		if (abs(iMSU1_Fallback_Enable - 1) < 0.5) then
		   VAR(iUSRCD1_Strt_Var + 6) = rPwf_limit_pu
		   CON(iPPC_P_Ctrl_CON + 10) = rPwf_limit_pu
		end if
C2C
        VAR(iUSRCD1_Strt_Var + 0) = VAR(iPPC_Q_Ctrl_VAR + iMSU1_Discharge_VAR)
        VAR(iUSRCD1_Strt_Var + 1) = VAR(iPPC_Q_Ctrl_VAR + iMSU2_Discharge_VAR)
        VAR(iUSRCD1_Strt_Var + 2) = VAR(iPPC_Q_Ctrl_VAR + iMSU3_Discharge_VAR)
        VAR(iUSRCD1_Strt_Var + 3) = VAR(iPPC_Q_Ctrl_VAR + iMSU4_Discharge_VAR)	
C2C
		if (abs(iFreq_Injection_Enable - 1) < 0.5) then
		   VAR(iUSRCD1_Strt_Var + 8)    = rTemp
		   STATE(iPPC_P_Ctrl_STATE + 2) = rTemp
		   STORE(iPPC_P_Ctrl_STATE + 2) = rTemp	
		end if
C2C
        Return
      End If
C2C		 
C++
C**************************** MODE > 4 *************************************************
C--
1000  IF (MODE==6) THEN
         Write (IPRT, 5018) iBusNbr, iNum_Icon,iNum_Con,iNum_State,iNum_Var 
     &,           (ICON(iIndex), iIndex = iUSRCD1_Strt_Icon, iUSRCD1_Strt_Icon  + iNum_Icon  - 1)
     &,           (CON(iIndex), iIndex = iUSRCD1_Strt_Con ,iUSRCD1_Strt_Con + iNum_Con   - 1)
         RETURN
      ENDIF 
C     MODE 5 ACTIVITY DOCU
      IF (MODE == 5) THEN
        CALL DOCUHD(*1900)
        chCON_ARRAY(1:Array_Character_Number)  = 'No Description'
        Write(IPRT,5001) iBusNbr
        Write(IPRT,5002) iUSRCD1_Strt_Con, iUSRCD1_Strt_Con   + iNum_Con   - 1,
     &                   iUSRCD1_Strt_Icon, iUSRCD1_Strt_Icon  + iNum_Icon  - 1,
     &                   iUSRCD1_Strt_State, iUSRCD1_Strt_State + iNum_State - 1,	 
     &                   iUSRCD1_Strt_Var, iUSRCD1_Strt_Var   + iNum_Var   - 1
C2C     ADDING PARAMETERS AND DESCRIPTION IN DOCU
             
        chCON_ARRAY	(	1	)	= 'PLimit_1111, pu'
        chCON_ARRAY	(	2	)	= 'PLimit_0111, pu'
        chCON_ARRAY	(	3	)	= 'PLimit_1011, pu'
        chCON_ARRAY	(	4	)	= 'PLimit_1101, pu'
        chCON_ARRAY	(	5	)	= 'PLimit_1110, pu'
        chCON_ARRAY	(	6	)	= 'PLimit_0011, pu'
        chCON_ARRAY	(	7	)	= 'PLimit_1001, pu'
        chCON_ARRAY	(	8	)	= 'PLimit_1100, pu'
        chCON_ARRAY	(	9	)	= 'PLimit_1000, pu'
        chCON_ARRAY	(	10	)	= 'PLimit_0100, pu'
        chCON_ARRAY	(	11	)	= 'PLimit_0010, pu'
        chCON_ARRAY	(	12	)	= 'PLimit_0001, pu'
        chCON_ARRAY	(	13	)	= 'PLimit_0000, pu'
        chCON_ARRAY	(	14	)	= 'Reserved'
        chCON_ARRAY	(	15	)	= 'Reserved'
        chCON_ARRAY	(	16	)	= 'Freq inject start time, s'
        chCON_ARRAY	(	17	)	= 'Freq inject end time, s'
        chCON_ARRAY	(	18	)	= 'ROCOF 1, pu/s'
        chCON_ARRAY	(	19	)	= 'ROCOF 2, pu/s'
        chCON_ARRAY	(	20	)	= 'Freq deviation inject up limit, pu'
		chCON_ARRAY	(	21	)	= 'Freq deviation inject low limit, pu'
		chCON_ARRAY	(	22	)	= 'Reserved'
		chCON_ARRAY	(	23	)	= 'Reserved'
		chCON_ARRAY	(	24	)	= 'Reserved'
		chCON_ARRAY	(	25	)	= 'Reserved'
		chCON_ARRAY	(	26	)	= 'Reserved'
		chCON_ARRAY	(	27	)	= 'Reserved'
		chCON_ARRAY	(	28	)	= 'Reserved'
		chCON_ARRAY	(	29	)	= 'Reserved'
		chCON_ARRAY	(	30	)	= 'Reserved'
		
C2C
        Do iIndex = 1, Array_Character_Number
		   If (chCON_ARRAY(iIndex).NE.'No Description') Then
		      Write(IPRT,5003) chCON_ARRAY(iIndex),CON(iUSRCD1_Strt_Con + iIndex - 1)
		   End If
        End Do
C2C
         RETURN  !MODE 5
      ENDIF
C
1900  RETURN
C++
C-------- FORMAT STATEMENTS -----------------------------------------------------------
C--
5001  Format (//6X,'** USRCD1 ** at bus ', I0)
5002  Format (/6X,'Uses CONs ', I0,'-',I0, '   ICONs ', I0,'-',I0,
     &     '   STATEs ', I0,'-',I0,'   VARs ', I0,'-',I0)
5003  Format (20X,A30,'  :  ',G11.4)
5018   FORMAT(I7,' ''USRBUS'' ',6X,' ''USRCD1''  504 0', 4(3X,I2),(/20(3X,I3)),
     &       4(/7(6X,G11.4)),(/2(6X,G11.4)),'/')
	 
C++     
C========================= END of VESTAS PPC USER CODE MODEL============================
C--
      END SUBROUTINE USRCD1      