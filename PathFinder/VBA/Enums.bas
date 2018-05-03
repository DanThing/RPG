option Explicit
OPtion Compare Text

' example url to get json content is
' https://raw.githubusercontent.com/DanThing/PSRD-Data/release/core_rulebook/feat/acrobatic.json
global Const GitHub as string = "https://raw.githubusercontent.com/"
Global Const DefUser as String = "DanThing"
Global Const PFSD as String = "PSRD-Data"


Public Enum Dice
	d4 = 4
	d6 = 6
	d8 = 8
	d10 = 10
	d12 = 12
	d20 = 20
	d100 = 100
End Enum

Public Enum Coin
	cp
	sp
	gp
	pp
End Enum
