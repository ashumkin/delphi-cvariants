#$Id: Makefile,v 1.3 2001/02/16 00:56:05 juanco Exp $
ROOT=../..
include $(ROOT)\\Rules.mak

registration:
	$(DCC) -B TestDUnitGUI.dpr
	$(DCC) -B TestDUnit.dpr
	$(BIN_DIR)\\TestDUnit.exe
