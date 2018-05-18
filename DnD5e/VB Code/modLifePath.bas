Attribute VB_Name = "modLifePath"
Option Explicit
Option Compare Text

Private Function getWeirdStuff() As String
Dim r As Long
r = Roll(12)
Select Case r
    Case 1: getWeirdStuff = "You were turned into a toad and remained in that form for " & Roll(4) & " weeks."
    Case 2: getWeirdStuff = "You were petrified and remained a stone statue for a time until someone freed you."
    Case 3: getWeirdStuff = "You were enslaved by a hag, a satyr, or some other being and lived in that creature's thrall for " & Roll(6) & " years."
    Case 4: getWeirdStuff = "A dragon held you as a prisoner for " & Roll(12) & " months until adventurers killed it."
    Case 5: getWeirdStuff = "You were taken captive and lived as a slave in the Underdark until you escaped."
    Case 6: getWeirdStuff = "You served a powerful adventurer as a hireling. You have only recently left that service."
    Case 7: getWeirdStuff = "You went insane for " & Roll(6) & " years and recently regained your sanity. A tic or some other bit of odd behavior might linger."
    Case 8: getWeirdStuff = "You discovered that a lover of yours was secretly a dragon."
    Case 9: getWeirdStuff = "You were captured by a cult and nearly sacrificed on an altar to the foul being the cultists served. You escaped, but you fear they will find you."
    Case 10: getWeirdStuff = "You met a demigod, an archdevil, an archfey, a demon lord, or a titan, and you lived to tell the tale."
    Case 11: getWeirdStuff = "You were swallowed by a giant fish and spent a month in its gullet before you escaped."
    Case 12: getWeirdStuff = "A powerful being granted you a wish, but you squandered it on something frivolous."
End Select
End Function

Private Function getWar() As String
Dim r As Long
r = Roll(12)
Select Case r
    Case 1: getWar = "You were knocked out and left for dead. You woke up hours later with no recollection of the battle."
    Case Is < 4: getWar = "You were badly injured in the fight, and you still bear the awful scars of those wounds."
    Case 4: getWar = "You ran away from the battle to save your life, but you still feel shame for your cowardice."
    Case Is < 8: getWar = "You suffered only minor injuries, and the wounds all healed without leaving scars."
    Case Is < 10: getWar = "You survived the battle, but you suffer from terrible nightmares in which you relive the experience."
    Case Is < 12: getWar = "You escaped the battle unscathed, though many of your friends were injured or lost."
    Case 12: getWar = "You acquitted yourself well in battle and are remembered as a hero. You might have received a medal for your bravery"
End Select
End Function

Private Function getTragedy() As String
Dim r As Long, d As Long
r = Roll(12)
Select Case r
    Case Is < 3: getTragedy = "A family member or a close friend died. Cause of death " & CauseofDeath
    Case 3: getTragedy = "A friendship ended bitterly, and the other person is now hostile to you. The cause might have been a misunderstanding or something you or the former friend did."
    Case 4: getTragedy = "You lost all your possessions in a disaster, and you had to rebuild your life."
    Case 5: getTragedy = "You were imprisoned for a crime you didn't commit and spent " & Roll(6) & " years at hard labor, in jail, or shackled to an oar in a slave galley."
    Case 6: getTragedy = "War ravaged your home community, reducing everything to rubble and ruin. in the aftermath, you either helped your town rebuild or moved somewhere else."
    Case 7: getTragedy = "A lover disappeared without a trace. You have been looking for that person ever since."
    Case 8: getTragedy = "A terrible blight in your home community caused crops to fail, and many starved. You lost a sibling or some other family member."
    Case 9: getTragedy = "You did something that brought terrible shame to you in the eyes of your family. You might have been involved in a scandal, dabbled in dark magic, or offended someone important. The attitude of your family members toward you becomes indifferent at best, though they might eventually forgive you."
    Case 10: getTragedy = "For a reason you were never told, you were exiled from your community. You then either wandered in the wilderness for a time or promptly found a new place to live."
    Case 11
        d = Roll(4)
        Select Case d
            Case Is < 3
                getTragedy = "A romantic relationship ended with bad feelings"
            Case Is < 5
                getTragedy = "A romantic relationship ended but it was amicable"
        End Select
    Case 12
        d = Roll(4)
        Select Case d
            Case Is < 3: getTragedy = "A current or prospective romantic partner of yours died. Cause of death " & CauseofDeath & ". It was your fault."
            Case Is < 5: getTragedy = "A current or prospective romantic partner of yours died. Cause of death " & CauseofDeath & "."
        End Select
    End Select
End Function

Private Function getSupernatural() As String
Dim r As Long
r = Roll(100)
Select Case r
    Case Is < 6: getSupernatural = "You were ensorcelled by a fey and enslaved for " & Roll(6) & " years before you escaped."
    Case Is < 11: getSupernatural = "You saw a demon and ran away before it could do anything to you."
    Case Is < 16: getSupernatural = "A devil tempted you. Make a DC 10 Wisdom saving throw. On a failed save, your alignment shifts one step toward evil (if it\'s not evil already), and you start the game with an additional ${random.dice('1d20') + 50} gp."
    Case Is < 21: getSupernatural = "You woke up one morning miles from your home, with no idea how you got there."
    Case Is < 31: getSupernatural = "You visited a holy site and felt the presence of the divine there."
    Case Is < 41: getSupernatural = "You witnessed a falling red star, a face appearing in the frost, or some other bizarre happening. You are certain that it was an omen of some sort."
    Case Is < 51: getSupernatural = "You escaped certain death and believe it was the intervention of a god that saved you."
    Case Is < 61: getSupernatural = "You witnessed a minor miracle."
    Case Is < 71: getSupernatural = "You explored an empty house and found it to be haunted."
    Case Is < 76: getSupernatural = "You were briefly possessed by a " & getPossessedBy
    Case Is < 81: getSupernatural = "You saw a ghost."
    Case Is < 86: getSupernatural = "You saw a ghoul feeding on a corpse."
    Case Is < 91: getSupernatural = "A celestial or a fiend visited you in your dreams to give a warning of dangers to come."
    Case Is < 96: getSupernatural = "You briefly visited the Feywild or the Shadowfell."
    Case Is > 96: getSupernatural = "You saw a portal that you believe leads to another plane of existence."
End Select
End Function

Private Function getPossessedBy() As String
Dim r As Long
r = Roll(6)
Select Case r
    Case 1: getPossessedBy = "celestial"
    Case 2: getPossessedBy = "devil"
    Case 3: getPossessedBy = "demon"
    Case 4: getPossessedBy = "fey"
    Case 5: getPossessedBy = "elemental"
    Case 6: getPossessedBy = "undead"
End Select
End Function

Private Function getCrime() As String
Dim r As Long
r = Roll(8)
Select Case r
    Case 1: getCrime = "murder"
    Case 2: getCrime = "theft"
    Case 3: getCrime = "burglary"
    Case 4: getCrime = "assault"
    Case 5: getCrime = "smuggling"
    Case 6: getCrime = "kidnapping"
    Case 7: getCrime = "extortion"
    Case 8: getCrime = "counterfeiting"
End Select
End Function

Private Function getPunishment() As String
Dim r As Long
r = Roll(12)
Select Case r
    Case Is < 4: getPunishment = "You did not commit the crime and were exonerated after being accused."
    Case Is < 7: getPunishment = "You committed the crime or helped do so, but nonetheless the authorities found you not guilty."
    Case Is < 9: getPunishment = "You were nearly caught in the act. You had to flee and are wanted in the community where the crime occurred."
    Case Is < 13: getPunishment = "You were caught and convicted. You spent time in jail, chained to an oar, or performing hard labor. You served a sentence of ${random.dice('1d4')} years or succeeded in escaping after that much time."
End Select
End Function

Private Function getBoon() As String
Dim r As Long
r = Roll(10)
Select Case r
    Case 1: getBoon = "A friendly wizard gives you a spell scroll containing one cantrip"
    Case 2: getBoon = "You save the life of a commoner, who now owes you a life debt. This individual accompanies you on your travels and performs mundane tasks for you, but will leave if neglected, abused, or imperiled."
    Case 3: getBoon = "You find a riding horse."
    Case 4: getBoon = "You find some money. You have " & Roll(20) & " gp in addition to your regular starting funds."
    Case 5: getBoon = "A relative bequeathed you a simple weapon of your choice."
    Case 6: getBoon = "You find " & getTrinket
    Case 7: getBoon = "You perform a service for a local temple. The next time you visit the temple, you can receive healing up to your hit point maximum."
    Case 8: getBoon = "A friendly alchemist gifts you with a potion of healing or a flask of acid, as you choose."
    Case 9: getBoon = "You find a treasure map."
    Case 10: getBoon = "A distant relative leaves you a stipend that enables you to live at the comfortable lifestyle for ${random.dice('1d20')} years. If you choose to live at a higher lifestyle, you reduce the price of the lifestyle by 2 gp during that time period."
    Case Else
        Err.Raise CustomError.Err9("Get Boon", "Unable to get boon.")
End Select
End Function

Private Function getEvent(preventLoop As Boolean) As String
StartOver:
Dim r As Long, d As Long
r = Roll(100)
Select Case r
    Case 1: getEvent = getWeirdStuff
    Case Is < 11: getTragedy
    Case Is < 21: getBoon
    Case Is < 31
        d = Roll(4)
        Select Case d
            Case 1: getEvent = " have a child"
            Case Is < 4: getEvent = " fall in love"
            Case 4: getEvent = " get married"
        End Select
    Case Is < 41
        d = Roll(4)
        Select Case d
            Case Is < 3: getEvent = "make an enemy of an adventuring " & getClass & ", but it wasn't your fault"
            Case Is < 5: getEvent = "make an enemy of an adventuring " & getClass & ", and it was your fault."
            Case Else: getEvent = "make an enemy of an adventuring " & getClass
        End Select
    Case Is < 51: getEvent = "make a friend of an adventuring " & getClass
    Case Is < 71: getEvent = "spend time working in a job related to your background. Start the game with an extra " & Roll(6, 2) & " gp."
    Case Is < 76: getEvent = "meet someone important."
    Case Is < 81: getEvent = "go on an adventure and " & getAdventure
    Case Is < 86: getEvent = getSupernatural
    Case Is < 91: getEvent = "fight in a battle. " & getWar
    Case Is < 96: getEvent = "are accused of " & getCrime & ". " & getPunishment
    Case Is < 100
        If preventLoop = False Then
            getEvent = getArcane(True)
        Else
            GoTo StartOver
        End If
    Case Else
        Err.Raise CustomError.Err9("Get Event", "Unable to get event.")
End Select
End Function

Function getArcane(preventLoop As Boolean) As String
Dim r As Long
StartOver:
r = Roll(10)
Select Case r
    Case 1: getArcane = "get charmed or frightened by a spell."
    Case 2: getArcane = "get injured by the effect of a spell."
    Case 3: getArcane = "witness a powerful spell being cast by a cleric, a druid, a sorcerer, a warlock, or a wizard."
    Case 4: getArcane = "drink a potion"
    Case 5: getArcane = "find a spell scroll and succeed in casting the Spell it contained."
    Case 6: getArcane = "are affected by teleportation magic."
    Case 7: getArcane = "turn invisible for a time."
    Case 8: getArcane = "identify an illusion for what it was."
    Case 9: getArcane = "see a creature being conjured by magic."
    Case 10
        If preventLoop = False Then
            getArcane = "Your fortune was read by a diviner. They told you: " & getEvent(True)
        Else
            GoTo StartOver
        End If
End Select
End Function


Function getAdventure() As String
Dim d As Long
    d = Roll(100)
    Select Case d
        Case Is <= 5: getAdventure = "come across a common magic item"
        Case Is <= 11: getAdventure = "nearly die. You have nasty scars on your body."
        Case Is <= 21: getAdventure = "suffer a grievous injury. Although the wound healed, it still pains you."
        Case Is <= 31: getAdventure = "are wounded, but in time you fully recovered."
        Case Is <= 41: getAdventure = "contract a disease while exploring a filthy warren. You will recover from the disease, but you have a persistent cough, pockmarks on your skin, or prematurely gray hair."
        Case Is <= 51: getAdventure = "are poisoned by a trap or a monster. You recovered, but the next time you must make a saving throw against poison, you make the saving throw with disadvantage."
        Case Is <= 61: getAdventure = "loose something of sentimental value to you during your adventure. Remove one trinket from your possessions."
        Case Is <= 71: getAdventure = "are terribly frightened by something you encountered and ran away, abandoning your companions to their fate."
        Case Is <= 81: getAdventure = "learn a great deal during your adventure. The next time you make an ability check or a saving throw, you have advantage on the roll."
        Case Is <= 91: getAdventure = "find some treasure on your adventure. You have " & Roll(6) & " gp left from your share of it."
        Case Is <= 100: getAdventure = "find a considerable amount of treasure on your adventure. You have " & Roll(20) + 50 & " gp left from your share of it."
        Case Else
            Err.Raise CustomError.Err9("Get Adventure", "Unable to get Adventure.")
    End Select
End Function

Function CauseofDeath() As String
Dim r As Long
r = Roll(12)
Select Case r
    Case 1: CauseofDeath = "Unknown"
    Case 2: CauseofDeath = "Murdered"
    Case 3: CauseofDeath = "Killed in battle"
    Case 4: CauseofDeath = "Accident related to class or occupation"
    Case 5: CauseofDeath = "Accident unrelated to class or occupation"
    Case Is < 7: CauseofDeath = "Natural causes, such as disease or old age"
    Case 8: CauseofDeath = "Apparent suicide"
    Case 9: CauseofDeath = "Torn apart by an animal or a natural disaster"
    Case 10: CauseofDeath = "Consumed by a monster"
    Case 11: CauseofDeath = "Executed for a crime or tortured to death"
    Case 12: CauseofDeath = "Bizarre event, such as being hit by a meteorite, struck down by an angry god, or killed by a hatching slaad egg"
End Select
End Function
