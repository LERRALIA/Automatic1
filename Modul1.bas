Attribute VB_Name = "Modul1"
Option Explicit
Public lngZeit As Long
Public lngEinzelPic As Long
Public Function Spruch_des_Tages() As String
    On Error GoTo LOKAL_ERROR
    
    Spruch_des_Tages = "Spruch des Tages:" & vbCrLf
    Dim yearday As Integer
    Dim i As Integer
    i = 271
    
    yearday = DateDiff("d", CDate("1/1/" & Year(Now)), DateValue(Now)) + 1
    
    Select Case yearday
        Case 1
            Spruch_des_Tages = Spruch_des_Tages & "Die Neigung des Menschen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "kleine Dinge für wichtig zu halten," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "hat sehr viel Großes hervorgebracht." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Georg Christoph Lichtenberg)" & vbCrLf
            
          Case 2
            
            Spruch_des_Tages = Spruch_des_Tages & "Man hat einen Menschen noch lange nicht bekehrt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wenn man ihn zum Schweigen gebracht hat." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(John Morley of Blackburn)" & vbCrLf
            
        Case 3
            
            Spruch_des_Tages = Spruch_des_Tages & "Wenn einer keine Angst hat," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "hat er keine Fantasie." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Erich Kästner)" & vbCrLf
            
         Case 4
            
            Spruch_des_Tages = Spruch_des_Tages & "Und wenn ich wüsste," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass morgen die Welt zugrunde ging," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "pflanzte ich doch heute noch einen Baum." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Martin Luther)" & vbCrLf
            
         Case 5
            
            Spruch_des_Tages = Spruch_des_Tages & "Es kommt nicht darauf an," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dem Leben mehr Jahre zu geben," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "sondern den Jahren mehr Leben." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Alexis Carrel)" & vbCrLf
            
         Case 6
            
            Spruch_des_Tages = Spruch_des_Tages & "Manche Hähne glauben," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass die Sonne ihretwegen aufgeht." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Theodor Fontane)" & vbCrLf
            
        Case 7
            
            Spruch_des_Tages = Spruch_des_Tages & "Man verliert die meiste Zeit damit," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass man Zeit gewinnen will." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(John Steinbeck)" & vbCrLf
            
        Case 8
            
            Spruch_des_Tages = Spruch_des_Tages & "Die Selbstzufriedenheit ist in Wahrheit das Höchste," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "was man erhoffen kann." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Baruch Spinoza)" & vbCrLf
            
        Case 9
            
            Spruch_des_Tages = Spruch_des_Tages & "Die Ideen sind nicht verantwortlich dafür," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "was die Menschen aus ihnen machen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Werner Heisenberg)" & vbCrLf
            
         Case 10
            
            Spruch_des_Tages = Spruch_des_Tages & "Erst wenn der letzte Baum gerodet," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "der letzte Fluss vergiftet," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "der letzte Fisch gefangen ist," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "werdet ihr feststellen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass man Geld nicht essen kann." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Indianisches Sprichwort)" & vbCrLf
            
         Case 11
            
            Spruch_des_Tages = Spruch_des_Tages & "Was du ins Ohr flüsterst," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wird tausend Meilen weit gehört." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Chinesisches Sprichwort)" & vbCrLf
            
         Case 12
            
            Spruch_des_Tages = Spruch_des_Tages & "Es gibt Diebe, die nicht bestraft werden" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "und dem Menschen doch das Kostbarste stehlen:" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die Zeit." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Napoleon I.)" & vbCrLf
            
         Case 13
            
            Spruch_des_Tages = Spruch_des_Tages & "Keine Zukunft vermag gutzumachen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "was du in der Gegenwart versäumst." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Albert Schweitzer)" & vbCrLf
            
          Case 14
            
            Spruch_des_Tages = Spruch_des_Tages & "Krankheiten überfallen den Menschen" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "nicht wie ein Blitz aus heiterem Himmel," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "sondern sind die Folgen fortgesetzter Fehler" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wider die Natur." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Hippokrates)" & vbCrLf
            
         Case 15
            
            Spruch_des_Tages = Spruch_des_Tages & "Lernen ist wie Rudern gegen den Strom:" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Sobald man aufhört, treibt man zurück." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Benjamin Britten)" & vbCrLf
            
         Case 16
            
            Spruch_des_Tages = Spruch_des_Tages & "Einen Fehler durch eine Lüge zu verdecken heißt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "einen Flecken durch ein Loch zu ersetzen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Aristoteles)" & vbCrLf
            
          Case 17
            
            Spruch_des_Tages = Spruch_des_Tages & "Kaum hat mal einer ein bissel was," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "gleich gibt es welche, die ärgert das." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Wilhelm Busch)" & vbCrLf
            
          Case 18
            
            Spruch_des_Tages = Spruch_des_Tages & "An je weniger Bedürfnisse wir uns gewöhnt haben," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "um so weniger Entbehrungen drohen uns." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Leo Tolstoi)" & vbCrLf

        Case 19

            Spruch_des_Tages = Spruch_des_Tages & "Das ist eines der wohl tragischsten Missverständnisse unserer Zeit:" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Wir glauben, wenn etwas zweifelsfrei als falsch bewiesen ist," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "müsse das Gegenteil richtig sein." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Salvador de Madariaga y Rojo)" & vbCrLf
    
    
        Case 20

            Spruch_des_Tages = Spruch_des_Tages & "Die Menschen stolpern nicht über Berge," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "sondern über Maulwurfshügel." & vbCrLf
            
    
    
        Case 21

            Spruch_des_Tages = Spruch_des_Tages & "Unser Kopf ist rund," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "damit das Denken die Richtung wechseln kann." & vbCrLf
            
        Case 22

            Spruch_des_Tages = Spruch_des_Tages & "Optimisten sind Menschen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die wissen wie schlecht die Welt ist;" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Pessimisten sind Menschen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die es täglich neu erleben müssen." & vbCrLf
            
        Case 23

            Spruch_des_Tages = Spruch_des_Tages & "Was man mit Gewalt gewinnt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "kann man nur mit Gewalt behalten." & vbCrLf

        Case 24

            Spruch_des_Tages = Spruch_des_Tages & "Intelligenz wird oft verwechselt" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "mit der Fähigkeit," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "seine Dummheit besser verbergen" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "zu können als andere Menschen." & vbCrLf
            
        Case 25

            Spruch_des_Tages = Spruch_des_Tages & "Unser Zeitalter ist stolz auf Maschinen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die denken," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "und misstrauisch gegen Menschen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die es versuchen." & vbCrLf

 
        Case 26

            Spruch_des_Tages = Spruch_des_Tages & "Oft kommt das Glück durch eine Tür herein," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "von der man gar nicht wusste," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass man sie offen gelassen hatte." & vbCrLf
            
        Case 27

            Spruch_des_Tages = Spruch_des_Tages & "Man muss manchmal von einem Menschen fortgehen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "um ihn zu finden." & vbCrLf
            
        Case 28

            Spruch_des_Tages = Spruch_des_Tages & "Ein Gentleman ist ein Mann," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "der wenigstens von Zeit zu Zeit so ist," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wie er immer sein sollte." & vbCrLf
            
        Case 29

            Spruch_des_Tages = Spruch_des_Tages & "Arm ist nicht der," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "der wenig hat," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "sondern der," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "der nicht genug bekommen kann." & vbCrLf
            
        Case 30

            Spruch_des_Tages = Spruch_des_Tages & "Ein Dummkopf findet immer einen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "der noch dümmer ist," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "der ihn bewundert." & vbCrLf
            
        Case 31

            Spruch_des_Tages = Spruch_des_Tages & "Genieße den heutigen Tag," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "denn mit dem heutigen Tag" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "beginnt der Rest deines Lebens." & vbCrLf
            
        Case 32

            Spruch_des_Tages = Spruch_des_Tages & "Willst du den Charakter eines Menschen erkennen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "so gib ihm Macht." & vbCrLf
            
        Case 33

            Spruch_des_Tages = Spruch_des_Tages & "Die ganze Kunst des Redens besteht darin," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "zu wissen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "was man nicht sagen darf." & vbCrLf
            
        Case 34

            Spruch_des_Tages = Spruch_des_Tages & "Oft erkennt man wie dumm man war," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "aber nie," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wie dumm man ist." & vbCrLf
            
            
        Case 35

            Spruch_des_Tages = Spruch_des_Tages & "Es kann mich niemand daran hindern," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "klüger zu werden." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Adenauer)" & vbCrLf
            
        Case 36

            Spruch_des_Tages = Spruch_des_Tages & "Wer glaubt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "ein Teamleiter leite ein Team," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Wer glaubt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "ein Zitronenfalter falte Zitronen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Team)" & vbCrLf
            
        Case 37

            Spruch_des_Tages = Spruch_des_Tages & "Ein freundlicher Blick," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "eine Geste der Zuneigung," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "gilt mehr als viele Worte." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Anna Strafinger)" & vbCrLf
            
        Case 38

            Spruch_des_Tages = Spruch_des_Tages & "Die Tat ist vergangen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die Denkmäler bleiben." & vbCrLf
            
        Case 39

            Spruch_des_Tages = Spruch_des_Tages & "Ich werde nie behaupten das ich die Nummer 1 bin," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "doch ich werde nie zugeben die Nummer 2 zu sein !" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Bruce Lee)" & vbCrLf
            
        Case 40

            Spruch_des_Tages = Spruch_des_Tages & "Beim Reden ist schätzenswert nicht die Zahl," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "sondern die Nützlichkeit der Worte." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Tscheng I.)" & vbCrLf
            
        Case 41

            Spruch_des_Tages = Spruch_des_Tages & "Das absolute Wissen führt zum Pessimismus:" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die Kunst ist das Heilmittel dagegen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Friedrich Nietzsche)" & vbCrLf
            
        
        Case 42

            Spruch_des_Tages = Spruch_des_Tages & "Die besten Dinge im Leben sind nicht die," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die man für Geld bekommt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Albert Einstein)" & vbCrLf
            
        Case 43

            Spruch_des_Tages = Spruch_des_Tages & "Lache nie über die Dummheit der anderen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Sie ist deine Chance." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Winston Churchill)" & vbCrLf
            
            
        Case 44

            Spruch_des_Tages = Spruch_des_Tages & "Die Freiheit des Menschen liegt nicht darin," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass er tun kann was er will," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "sondern," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass er nicht tun muss, was er nicht will." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Jean-Jacques Rousseau)" & vbCrLf
            
            
        Case 45

            Spruch_des_Tages = Spruch_des_Tages & "Wer die Freiheit aufgibt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "um Sicherheit zu gewinnen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wird am Ende beides verlieren." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Benjamin Franklin)" & vbCrLf
            
        Case 46

            Spruch_des_Tages = Spruch_des_Tages & "Der Mensch ist erst wirklich tot," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wenn niemand mehr an ihn denkt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Bertolt Brecht)" & vbCrLf
            

        Case 47

            Spruch_des_Tages = Spruch_des_Tages & "Unser größter Ruhm ist nicht," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "niemals zu fallen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "sondern jedes Mal wieder aufzustehen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Ralph Waldo Emerson)" & vbCrLf
            
        Case 48

            Spruch_des_Tages = Spruch_des_Tages & "Der Freund ist einer," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "der alles von dir weiß," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "und der dich trotzdem liebt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Elbert Hubbard)" & vbCrLf
            
        Case 49

            Spruch_des_Tages = Spruch_des_Tages & "Der einzige Zwang des Lebens" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "sollte der Tod sein." & vbCrLf
            
        Case 50

            Spruch_des_Tages = Spruch_des_Tages & "Der Neid ist die aufrichtigste Form der Anerkennung." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Wilhelm Busch)" & vbCrLf
            
        Case 51

            Spruch_des_Tages = Spruch_des_Tages & "Reich wird man erst durch Dinge," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die man nicht begehrt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Mahatma Gandhi)" & vbCrLf
            
        Case 52
        
            Spruch_des_Tages = Spruch_des_Tages & "Hoffnung ist ein gutes Frühstück," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "aber ein schlechtes Abendbrot." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Francis Bacon)" & vbCrLf

        Case 53

            Spruch_des_Tages = Spruch_des_Tages & "Ich bin nicht sicher," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "mit welchen Waffen der dritte Weltkrieg ausgetragen wird," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "aber im vierten Weltkrieg werden sie mit Stöcken und Steinen kämpfen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Albert Einstein)" & vbCrLf
            
        Case 54

            Spruch_des_Tages = Spruch_des_Tages & "Der Letzte der lacht," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "ist der Erste der weint." & vbCrLf
        
        Case 55

            Spruch_des_Tages = Spruch_des_Tages & "Freundlichkeit ist eine Sprache," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die Taube hören" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "und Blinde lesen können." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Mark Twain)" & vbCrLf
            
        Case 56 'sa

            Spruch_des_Tages = Spruch_des_Tages & "Die Gestalterin findet für jedes Problem eine Lösung" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Die Verliererin findet für jede Lösung ein Problem" & vbCrLf
            
        Case 57

            Spruch_des_Tages = Spruch_des_Tages & "Wer glaubt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "etwas zu sein," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "hat aufgehört," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "etwas zu werden." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Philip Rosenthal)" & vbCrLf
            
        Case 58

            Spruch_des_Tages = Spruch_des_Tages & "Ein Freund ist der," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "der dich an die Melodie deines Herzens erinnert," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wenn du sie vergessen hast." & vbCrLf
            
        Case 59

            Spruch_des_Tages = Spruch_des_Tages & "Es gibt keine großen Entdeckungen und Fortschritte," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "solange es noch ein unglückliches Kind auf Erden gibt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Albert Einstein)" & vbCrLf
            
        Case 60

            Spruch_des_Tages = Spruch_des_Tages & "Es ist besser," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "ein einziges kleines Licht anzuzünden," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "als die Dunkelheit zu verfluchen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Konfuzius)" & vbCrLf
            
        Case 61

            Spruch_des_Tages = Spruch_des_Tages & "Vergib deinen Feinden," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "aber vergiss niemals ihre Namen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(John F.Kennedy)" & vbCrLf
            
        Case 62

            Spruch_des_Tages = Spruch_des_Tages & "Es gehört nur ein wenig Mut dazu," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "nicht das zu tun," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "was alle tun." & vbCrLf
            
        Case 63 'sa
        
            Spruch_des_Tages = Spruch_des_Tages & "Die Gestalterin sagt: 'Es mag schwierig sein, aber es ist möglich'" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Die Verliererin sagt: 'Es ist möglich, aber zu schwierig'" & vbCrLf

        Case 64

            Spruch_des_Tages = Spruch_des_Tages & "Wer A sagt, muss nicht B sagen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Er kann auch erkennen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass A falsch war." & vbCrLf
            
        Case 65

            Spruch_des_Tages = Spruch_des_Tages & "Die Hälfte aller Menschen wollen abnehmen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die andere Hälfte verhungert." & vbCrLf
            
        Case 66

            Spruch_des_Tages = Spruch_des_Tages & "Vielleicht ist es den Affen gar nicht recht," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass wir mit ihnen verwandt sind." & vbCrLf
            
        Case 67

            Spruch_des_Tages = Spruch_des_Tages & "Tradition heißt nicht," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die Asche zu bewahren," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "sondern das Feuer weiterzureichen." & vbCrLf
            
        Case 68

            Spruch_des_Tages = Spruch_des_Tages & "Der Mensch hat die Atombombe erfunden." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Keine Maus der Welt wäre je auf die Idee gekommen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "eine Mausefalle zu bauen." & vbCrLf
            
        Case 69
            
            Spruch_des_Tages = Spruch_des_Tages & "Solange Menschen denken," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass Tiere nicht fühlen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "müssen Tiere fühlen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass Menschen nicht denken!" & vbCrLf
            
        Case 70 'sa
        
            Spruch_des_Tages = Spruch_des_Tages & "Die Gestalterin hat immer einen Plan" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Die Verliererin hat immer eine Ausrede" & vbCrLf
            
        Case 71
        
            Spruch_des_Tages = Spruch_des_Tages & "Es ist von großem Vorteil, die Fehler," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "aus denen man lernen kann, recht früh zu machen." & vbCrLf
        
        Case 72
        
            Spruch_des_Tages = Spruch_des_Tages & "Geizhälse sind unangenehme Zeitgenossen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "aber angenehme Vorfahren." & vbCrLf
            
        Case 73
        
            Spruch_des_Tages = Spruch_des_Tages & "Wen wir lieben," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dem geben wir die Macht uns Leiden zu bereiten." & vbCrLf
        
        Case 74
        
            Spruch_des_Tages = Spruch_des_Tages & "Beginne Deinen Tag mit einem Lächeln!" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Deine Mitmenschen werden es Dir danken!" & vbCrLf
        
        Case 75
        
            Spruch_des_Tages = Spruch_des_Tages & "Das Recht auf Dummheit gehört zur Garantie" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "der freien Entfaltung der Persönlichkeit." & vbCrLf
        
        Case 76
        
            Spruch_des_Tages = Spruch_des_Tages & "Ich bin nicht geboren," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "um so zu handeln wie andere mich haben wollen." & vbCrLf
            
        Case 77
        
            Spruch_des_Tages = Spruch_des_Tages & "Man sieht nur mit dem Herzen gut," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "das Wesentliche ist für die Augen unsichtbar." & vbCrLf
        
        Case 78
        
            Spruch_des_Tages = Spruch_des_Tages & "Der Verstand ist wie eine Fahrkarte." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Sie hat nur dann einen Sinn wenn sie benutzt wird." & vbCrLf
            
        Case 79
        
            Spruch_des_Tages = Spruch_des_Tages & "Wir brauchen Kinder nicht erziehen -" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die machen uns eh alles nach." & vbCrLf
        
        Case 80
        
            Spruch_des_Tages = Spruch_des_Tages & "Phantasie ist viel wichtiger als Wissen, " & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "denn Wissen ist begrenzt." & vbCrLf
            
        Case 81
        
            Spruch_des_Tages = Spruch_des_Tages & "Wir sind nicht nur verantwortlich für das," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "was wir tun," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "sondern auch für das," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "was wir nicht tun." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Molière)" & vbCrLf
    
        Case 82
        
            Spruch_des_Tages = Spruch_des_Tages & "Liebe ist die stärkste Macht der Welt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "und doch ist sie die demütigste," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die man sich vorstellen kann." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Mahatma Gandhi)" & vbCrLf
    
        Case 83
        
            Spruch_des_Tages = Spruch_des_Tages & "Sei du selbst die Veränderung," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die du dir wünschst für diese Welt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Mahatma Gandhi)" & vbCrLf
            
        Case 84
        
            Spruch_des_Tages = Spruch_des_Tages & "Den größten Fehler," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "den man im Leben machen kann," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "immer Angst zu haben," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "einen Fehler zu machen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Dietrich Bonhoeffer)" & vbCrLf
            
        Case 85
        
            Spruch_des_Tages = Spruch_des_Tages & "Es ist nicht zu wenig Zeit," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die wir haben," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "sondern es ist zu viel Zeit," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die wir nicht nutzen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Lucius Annaeus Seneca)" & vbCrLf

        Case 86
        
            Spruch_des_Tages = Spruch_des_Tages & "Die beste und sicherste Tarnung ist" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "immer noch die blanke und nackte Wahrheit." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Die glaubt niemand!" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Max Frisch)" & vbCrLf
            
        Case 87
        
            Spruch_des_Tages = Spruch_des_Tages & "Müde macht uns die Arbeit," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die wir liegenlassen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "nicht die, die wir tun." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Marie von Ebner-Eschenbach)" & vbCrLf
            
        Case 88
        
            Spruch_des_Tages = Spruch_des_Tages & "Neid ist die Eifersucht darüber," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass sich Gott auch mit" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "anderen Menschen außer uns beschäftigt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Ernst R.Hauschka)" & vbCrLf

        Case 89
        
            Spruch_des_Tages = Spruch_des_Tages & "Teamarbeit ist, wenn vier Leute für eine Arbeit bezahlt werden," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die drei besser machen könnten," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wenn sie nur zu zweit gewesen wären und einer davon krank zu Bett läge." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Martin Wolgast)" & vbCrLf
            
        Case 90
        
            Spruch_des_Tages = Spruch_des_Tages & "Weil Denken die schwerste Arbeit ist," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die es gibt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "beschäftigen sich auch nur wenige damit." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Henry Ford)" & vbCrLf
            
        Case 91
        
            Spruch_des_Tages = Spruch_des_Tages & "Je mehr Vergnügen du an deiner Arbeit hast," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "um so besser wird sie bezahlt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Mark Twain)" & vbCrLf
            
        Case 92
        
            Spruch_des_Tages = Spruch_des_Tages & "Ein Faulpelz ist ein Mensch," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "der sich keine Arbeit damit macht," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "sein Nichtstun zu begründen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Gabriel Laub)" & vbCrLf
            
        Case 93
        
            Spruch_des_Tages = Spruch_des_Tages & "Versuchung ist ein Parfum," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "das man so lange riecht," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "bis man die Flasche haben möchte." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Jean-Paul Belmondo)" & vbCrLf

        Case 94
        
            Spruch_des_Tages = Spruch_des_Tages & "Düfte sind die Gefühle der Blumen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Heinrich Heine)" & vbCrLf
            
        Case 95
        
            Spruch_des_Tages = Spruch_des_Tages & "Faulheit: der Hang zur Ruhe" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "ohne vorhergehende Arbeit." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Immanuel Kant)" & vbCrLf
            
        Case 96
        
            Spruch_des_Tages = Spruch_des_Tages & "Wenn ein Mann eine Frau" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "nicht mehr riechen kann," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "hilft auch das beste Parfüm nichts mehr." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Helen Vita)" & vbCrLf
            
        Case 97
        
            Spruch_des_Tages = Spruch_des_Tages & "Persönlichkeiten werden nicht durch schöne Reden geformt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "sondern durch Arbeit und eigene Leistung." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Albert Einstein)" & vbCrLf
            
        Case 98
        
            Spruch_des_Tages = Spruch_des_Tages & "Glück ist ein Parfüm," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "das du nicht auf andere sprühen kannst," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "ohne selbst ein paar Tropfen abzubekommen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Ralph Waldo Emerson)" & vbCrLf
            
        Case 99
        
            Spruch_des_Tages = Spruch_des_Tages & "Faulheit ist die Furcht" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "vor bevorstehender Arbeit." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Marcus Tullius Cicero)" & vbCrLf
            
        Case 100
        
            Spruch_des_Tages = Spruch_des_Tages & "Die Arbeit ist etwas Unnatürliches." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Die Faulheit allein ist göttlich." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Anatole France)" & vbCrLf
            
        Case 101
        
            Spruch_des_Tages = Spruch_des_Tages & "Zwei Dinge sind zu unserer Arbeit nötig:" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Unermüdliche Ausdauer und die Bereitschaft," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "etwas, in das man viel Zeit und Arbeit gesteckt hat," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wieder wegzuwerfen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Albert Einstein)" & vbCrLf
            
        Case 102
        
            Spruch_des_Tages = Spruch_des_Tages & "Eine Frau ohne Geheimnisse" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "ist wie eine Blume ohne Duft." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Maurice Chevalier)" & vbCrLf

        Case 103
        
            Spruch_des_Tages = Spruch_des_Tages & "Es gibt eine Zeit für die Arbeit." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Und es gibt eine Zeit für die Liebe." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Mehr Zeit hat man nicht." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Coco Chanel)" & vbCrLf

        Case 104
        
            Spruch_des_Tages = Spruch_des_Tages & "Glück ist Liebe," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "nichts anderes." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Wer lieben kann," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "ist glücklich." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Hermann Hesse)" & vbCrLf
            
        Case 105
        
            Spruch_des_Tages = Spruch_des_Tages & "Wenn die Frauen verblühen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "verduften die Männer." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Heinrich Zille)" & vbCrLf
            
        Case 106
        
            Spruch_des_Tages = Spruch_des_Tages & "Wenn man glücklich ist," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "soll man nicht noch glücklicher sein wollen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Theodor Fontane)" & vbCrLf
            
        Case 107
        
            Spruch_des_Tages = Spruch_des_Tages & "Freude an der Arbeit lässt das Werk trefflich geraten." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Aristoteles)" & vbCrLf

        Case 108
        
            Spruch_des_Tages = Spruch_des_Tages & "Das Geheimnis des Glücks liegt nicht im Besitz," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "sondern im Geben." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Wer andere glücklich macht, wird glücklich." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(André Gide)" & vbCrLf
            
        Case 109
        
            Spruch_des_Tages = Spruch_des_Tages & "Für seine Arbeit muss man Zustimmung suchen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "aber niemals Beifall." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Charles-Louis de Montesquieu)" & vbCrLf
            
        Case 110
        
            Spruch_des_Tages = Spruch_des_Tages & "Man will nicht nur glücklich sein," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "sondern glücklicher als die anderen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Und das ist deshalb so schwer," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "weil wir die anderen für glücklicher halten, als sie sind." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Charles-Louis de Montesquieu)" & vbCrLf
            
        Case 111
        
            Spruch_des_Tages = Spruch_des_Tages & "Die Neider sterben wohl," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "doch niemals stirbt der Neid." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Molière)" & vbCrLf
            
        Case 112
        
            Spruch_des_Tages = Spruch_des_Tages & "Arbeit ist einer der besten Erzieher des Charakters." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Samuel Smiles)" & vbCrLf
            
        Case 113
        
            Spruch_des_Tages = Spruch_des_Tages & "Wenn man ganz bewusst acht Stunden täglich arbeitet," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "kann man es dazu bringen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Chef zu werden und vierzehn Stunden täglich zu arbeiten." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Robert Frost)" & vbCrLf

        Case 114
        
            Spruch_des_Tages = Spruch_des_Tages & "Erfolg ist nur halb so schön," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wenn es niemanden gibt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "der einen beneidet." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Norman Mailer)" & vbCrLf
            
        Case 115
        
            Spruch_des_Tages = Spruch_des_Tages & "Arbeit um der Arbeit willen" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "ist gegen die menschliche Natur." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(John Locke)" & vbCrLf
            
        Case 116
        
            Spruch_des_Tages = Spruch_des_Tages & "Glück bedeutet eine gute Gesundheit" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "und ein schlechtes Gedächtnis." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Ingrid Bergman)" & vbCrLf
            
        Case 117
        
            Spruch_des_Tages = Spruch_des_Tages & "Arbeit ist das beste Mittel gegen Verzweiflung." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Sir Arthur Conan Doyle)" & vbCrLf
            
         Case 118
        
            Spruch_des_Tages = Spruch_des_Tages & "Gegenüber der Fähigkeit," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die Arbeit eines einzigen Tages sinnvoll zu ordnen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "ist alles andere im Leben ein Kinderspiel." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Johann Wolfgang von Goethe)" & vbCrLf
            
        Case 119
        
            Spruch_des_Tages = Spruch_des_Tages & "Mitleid bekommt man geschenkt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Neid muss man sich verdienen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Robert Lembke)" & vbCrLf
            
         Case 120
        
            Spruch_des_Tages = Spruch_des_Tages & "Das Vergleichen ist das Ende des Glücks" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "und der Anfang der Unzufriedenheit." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Søren Kierkegaard)" & vbCrLf
            
        Case 121
        
            Spruch_des_Tages = Spruch_des_Tages & "Wenn ein Mann will," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass seine Frau zuhört," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "brauch er nur mit einer anderen zu reden." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Liza Minnelli)" & vbCrLf

        Case 122
        
            Spruch_des_Tages = Spruch_des_Tages & "Eine Ehe ist wie ein Restaurantbesuch" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "- man denkt immer," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "man hat das Beste gewählt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "bis man sieht, was der Nachbar bekommt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Bernd Stelter)" & vbCrLf
            
        Case 123
        
            Spruch_des_Tages = Spruch_des_Tages & "Die Anzahl unserer Neider" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "bestätigt unsere Fähigkeiten." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Oscar Wilde)" & vbCrLf
            
        Case 124
        
            Spruch_des_Tages = Spruch_des_Tages & "Glück entsteht oft durch Aufmerksamkeit" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "in kleinen Dingen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Unglück oft durch Vernachlässigung kleiner Dinge." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Wilhelm Busch)" & vbCrLf
            
        Case 125
        
            Spruch_des_Tages = Spruch_des_Tages & "Mancher lehnt eine gute Idee" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "bloß deshalb ab," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "weil sie nicht von ihm ist." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Luis Buñuel)" & vbCrLf
            
        Case 126
        
            Spruch_des_Tages = Spruch_des_Tages & "Es stimmt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass Geld nicht glücklich macht." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Allerdings meint man damit das Geld der anderen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(George Bernard Shaw)" & vbCrLf
            
        Case 127
        
            Spruch_des_Tages = Spruch_des_Tages & "Wenn man erfolgreich ist," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dann überschlagen sich die Freunde," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "aber erst wenn man einen Misserfolg hat," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dann freuen sie sich wirklich." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Harry S.Truman)" & vbCrLf
            
        Case 128

            Spruch_des_Tages = Spruch_des_Tages & "Gutes kann niemals aus Lüge und Gewalt entstehen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Mahatma Gandhi)" & vbCrLf
            
        Case 129

            Spruch_des_Tages = Spruch_des_Tages & "Es ist nicht genug zu wissen" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "- man muss auch anwenden." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Es ist nicht genug zu wollen" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "- man muss auch tun." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Johann Wolfgang von Goethe)" & vbCrLf
            
        Case 130

            Spruch_des_Tages = Spruch_des_Tages & "Nenne dich nicht arm," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "weil deine Träume nicht in Erfüllung gegangen sind;" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wirklich arm ist nur," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "der nie geträumt hat." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Marie von Ebner-Eschenbach)" & vbCrLf
            
        Case 131

            Spruch_des_Tages = Spruch_des_Tages & "Es zählt immer allein," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "was Du tust," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "nicht, was Du anderen zu tun empfiehlst!" & vbCrLf
            
'***
        Case 132

            Spruch_des_Tages = Spruch_des_Tages & "Die Welt ist so schön und wert," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass man um sie kämpft." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Ernest Hemingway)" & vbCrLf
            
        Case 133

            Spruch_des_Tages = Spruch_des_Tages & "Alle sind Irre;" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "aber wer seinen Wahn zu analysieren versteht," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wird Philosoph genannt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Ambrose Bierce)" & vbCrLf
            
        Case 134

            Spruch_des_Tages = Spruch_des_Tages & "Einen sicheren Freund" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "erkennt man in unsicherer Sache." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Marcus Tullius Cicero)" & vbCrLf
        Case 135

            Spruch_des_Tages = Spruch_des_Tages & "Mit einer geballten Faust" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "kann man keinen Händedruck wechseln." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Indira Gandhi)" & vbCrLf
            
        Case 136

            Spruch_des_Tages = Spruch_des_Tages & "Viele Menschen sind gut erzogen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "um nicht mit vollem Mund zu sprechen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "aber sie haben keine Bedenken," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "es mit leerem Kopf zu tun." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Orson Welles)" & vbCrLf


        Case 137

            Spruch_des_Tages = Spruch_des_Tages & "Wissenschaft ohne Religion ist lahm," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Religion ohne Wissenschaft ist blind." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Albert Einstein)" & vbCrLf
        Case 138

            Spruch_des_Tages = Spruch_des_Tages & "Niemand urteilt schärfer als der Ungebildete," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "er kennt weder Gründe noch Gegengründe." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Anselm Feuerbach)" & vbCrLf

        Case 139

            Spruch_des_Tages = Spruch_des_Tages & "Schlagfertigkeit ist etwas," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "worauf man erst 24 Stunden später kommt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Mark Twain)" & vbCrLf

        Case 140

            Spruch_des_Tages = Spruch_des_Tages & "Jeder sieht," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "was du scheinst." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Nur wenige fühlen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wie du bist." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Niccolò Machiavelli)" & vbCrLf

        Case 141

            Spruch_des_Tages = Spruch_des_Tages & "Ein Feigling ist ein Mensch," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "bei dem der Selbsterhaltungstrieb normal funktioniert." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Ambrose Bierce)" & vbCrLf


        Case 142

            Spruch_des_Tages = Spruch_des_Tages & "Eine schöne Uhr zeigt die Zeit an," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "eine schöne Frau lässt sie vergessen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Maurice Chevalier)" & vbCrLf

        Case 143

            Spruch_des_Tages = Spruch_des_Tages & "Das Geheimnis des Erfolges ist," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "den Standpunkt des anderen zu verstehen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Henry Ford)" & vbCrLf

        Case 144

            Spruch_des_Tages = Spruch_des_Tages & "Ein Mensch würde nie dazu kommen, etwas zu tun," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wenn er stets warten würde," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "bis er es so gut kann," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass niemand mehr einen Fehler entdecken könnte." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(John Henry Newman)" & vbCrLf

        Case 145

            Spruch_des_Tages = Spruch_des_Tages & "Ein erfolgreicher Mann ist ein Mann," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "der mehr verdient, als seine Frau ausgeben kann." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Eine erfolgreiche Frau ist eine," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die so einen Mann findet." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Mario Adorf)" & vbCrLf

        Case 146

            Spruch_des_Tages = Spruch_des_Tages & "Zyniker: ein Mensch, der die Dinge so sieht," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wie sie sind," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "und nicht, wie sie sein sollten." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Ambrose Bierce)" & vbCrLf


        Case 147

            Spruch_des_Tages = Spruch_des_Tages & "Alles ist gut. Der Mensch ist unglücklich," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "weil er nicht weiß, dass er glücklich ist." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Nur deshalb. Das ist alles, alles!" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Wer das erkennt, der wird gleich glücklich sein, sofort im selben Augenblick." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Fjodor Michailowitsch Dostojewski)" & vbCrLf

        Case 148

            Spruch_des_Tages = Spruch_des_Tages & "Man braucht zwei Jahre um sprechen zu lernen" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "und fünfzig, um schweigen zu lernen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Ernest Hemingway)" & vbCrLf


        Case 149

            Spruch_des_Tages = Spruch_des_Tages & "Die meisten großen Taten," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die meisten großen Gedanken" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "haben einen belächelnswerten Anfang." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Albert Camus)" & vbCrLf

'********

        Case 150

            Spruch_des_Tages = Spruch_des_Tages & "Vergessen können ist das Geheimnis ewiger Jugend." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Wir werden alt durch Erinnerung." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Erich Maria Remarque)" & vbCrLf

        Case 151

            Spruch_des_Tages = Spruch_des_Tages & "Der Langsamste," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "der sein Ziel nicht aus den Augen verliert," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "geht noch immer geschwinder," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "als jener, der ohne Ziel umherirrt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Gotthold Ephraim Lessing)" & vbCrLf

        Case 152

            Spruch_des_Tages = Spruch_des_Tages & "Die Ewigkeit dauert lange," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "besonders gegen Ende." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Woody Allen)" & vbCrLf

        Case 153

            Spruch_des_Tages = Spruch_des_Tages & "Nur die Weisen sind im Besitz von Ideen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Die meisten Menschen sind von Ideen besessen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Samuel Coleridge)" & vbCrLf

        Case 154

            Spruch_des_Tages = Spruch_des_Tages & "Auch wenn alle einer Meinung sind," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "können alle Unrecht haben." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Bertrand Russell)" & vbCrLf

        Case 155

            Spruch_des_Tages = Spruch_des_Tages & "Wer nicht mehr liebt " & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "und nicht mehr irrt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "der lasse sich begraben." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Johann Wolfgang von Goethe)" & vbCrLf

        Case 156

            Spruch_des_Tages = Spruch_des_Tages & "Für den gläubigen Menschen steht Gott am Anfang," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "für den Wissenschaftler am Ende aller seiner Überlegungen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Max Planck)" & vbCrLf

        Case 157

            Spruch_des_Tages = Spruch_des_Tages & "Die Jagd nach dem Sündenbock ist die einfachste." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Dwight D.Eisenhower)" & vbCrLf

        Case 158
        
            
            Spruch_des_Tages = Spruch_des_Tages & "Es gibt nichts Schöneres," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "als geliebt zu werden," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "geliebt um seiner selbst willen oder vielmehr trotz seiner selbst." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Victor Hugo)" & vbCrLf

        Case 159
        
            

            Spruch_des_Tages = Spruch_des_Tages & "Jeder junge Mensch macht früher oder später die verblüffende Entdeckung," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass auch Eltern gelegentlich Recht haben können." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(André Malraux)" & vbCrLf

        Case 160
        
            

            Spruch_des_Tages = Spruch_des_Tages & "Die Lüge ist wie ein Schneeball:" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Je länger man ihn wälzt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "desto größer wird er." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Martin Luther)" & vbCrLf

        Case 161
        
            

            Spruch_des_Tages = Spruch_des_Tages & "Nichts in der Welt ist so ansteckend" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wie Gelächter und gute Laune." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Charles Dickens)" & vbCrLf

        Case 162
        
            

            Spruch_des_Tages = Spruch_des_Tages & "Wer etwas Großes will," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "der muss sich zu beschränken wissen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wer dagegen alles will," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "der will in der Tat nichts und bringt es zu nichts." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Georg Wilhelm Friedrich Hegel)" & vbCrLf

        Case 163
        
            
            Spruch_des_Tages = Spruch_des_Tages & "Man wird in der Regel keinen Freund dadurch verlieren," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass man ihm ein Darlehen abschlägt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "aber sehr leicht dadurch," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass man es ihm gibt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Arthur Schopenhauer)" & vbCrLf

        Case 164
        
            

            Spruch_des_Tages = Spruch_des_Tages & "Kinder, die man nicht liebt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "werden Erwachsene," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die nicht lieben." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Pearl S.Buck)" & vbCrLf

        Case 165

            Spruch_des_Tages = Spruch_des_Tages & "Willst du dich am Ganzen erquicken," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "so musst du das Ganze im Kleinsten erblicken." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Johann Wolfgang von Goethe)" & vbCrLf

        Case 166

            Spruch_des_Tages = Spruch_des_Tages & "Warum die Hölle im Jenseits suchen?" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Sie ist schon im Diesseits vorhanden," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "im Herzen der Bösen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Jean-Jacques Rousseau)" & vbCrLf

        Case 167

            Spruch_des_Tages = Spruch_des_Tages & "Es gibt nichts Stilleres" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "als eine geladene Kanone." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Heinrich Heine)" & vbCrLf

        Case 168

            Spruch_des_Tages = Spruch_des_Tages & "Nicht die Vollkommenen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "sondern die Unvollkommenen brauchen unsere Liebe." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Oscar Wilde)" & vbCrLf

        Case 169

            Spruch_des_Tages = Spruch_des_Tages & "Jetzt sind die guten alten Zeiten," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "nach denen wir uns in zehn Jahren zurücksehnen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Peter Ustinov)" & vbCrLf

        Case 170

            Spruch_des_Tages = Spruch_des_Tages & "Im Entwurf, da zeigt sich das Talent," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "in der Ausführung die Kunst." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Marie von Ebner-Eschenbach)" & vbCrLf

        Case 171

            Spruch_des_Tages = Spruch_des_Tages & "Wenn Sie in der Politik etwas gesagt haben wollen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wenden Sie sich an einen Mann." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Wenn Sie etwas getan haben wollen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wenden Sie sich an eine Frau." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Margaret Thatcher)" & vbCrLf

        Case 172

            Spruch_des_Tages = Spruch_des_Tages & "Erfahrungen vererben sich nicht -" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "jeder muss sie allein machen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Kurt Tucholsky)" & vbCrLf

        Case 173

            Spruch_des_Tages = Spruch_des_Tages & "Ich fühle mich nicht zu dem Glauben verpflichtet," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass derselbe Gott, der uns mit Sinnen, Vernunft und Verstand ausgestattet hat," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "von uns verlangt, dieselben nicht zu benutzen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Galileo Galilei)" & vbCrLf

        Case 174

            Spruch_des_Tages = Spruch_des_Tages & "Erotik ist die Überwindung von Hindernissen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Das verlockendste und populärste Hindernis ist die Moral." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Karl Kraus)" & vbCrLf

        Case 175

            Spruch_des_Tages = Spruch_des_Tages & "Wenn du eine weise Antwort verlangst," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "musst du vernünftig fragen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Johann Wolfgang von Goethe)" & vbCrLf

        Case 176

            Spruch_des_Tages = Spruch_des_Tages & "Fallen ist weder gefährlich noch eine Schande." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Liegenbleiben ist beides." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Konrad Adenauer)" & vbCrLf

        Case 177

            Spruch_des_Tages = Spruch_des_Tages & "Zum Mitleiden gab die Natur vielen ein Talent," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "zur Mitfreude nur wenigen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Friedrich Hebbel)" & vbCrLf

        Case 178

            Spruch_des_Tages = Spruch_des_Tages & "Die Tränen lassen nichts gewinnen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wer schaffen will," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "muss fröhlich sein." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Theodor Fontane)" & vbCrLf

        Case 179

            Spruch_des_Tages = Spruch_des_Tages & "Das große Karthago führte drei Kriege." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Nach dem ersten war es noch mächtig." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Nach dem zweiten war es noch bewohnbar." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Nach dem dritten war es nicht mehr zu finden." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Bertolt Brecht)" & vbCrLf

        Case 180

            Spruch_des_Tages = Spruch_des_Tages & "Der Freund ist einer," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "der alles von dir weiß," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "und der dich trotzdem liebt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Elbert Hubbard)" & vbCrLf
            
        Case 181

            Spruch_des_Tages = Spruch_des_Tages & "Die meisten Probleme entstehen bei ihrer Lösung." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Leonardo da Vinci)" & vbCrLf

        Case 182

            Spruch_des_Tages = Spruch_des_Tages & "Wer sich zum Wurm macht," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "soll nicht klagen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wenn er getreten wird." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Immanuel Kant)" & vbCrLf

        Case 183

            Spruch_des_Tages = Spruch_des_Tages & "Zwischen Hochmut und Demut steht ein drittes," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dem das Leben gehört," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "und das ist der Mut." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Theodor Fontane)" & vbCrLf

        Case 184

            Spruch_des_Tages = Spruch_des_Tages & "Wenn man im Mittelpunkt einer Party stehen will," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "darf man nicht hingehen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Audrey Hepburn)" & vbCrLf

        Case 185

            Spruch_des_Tages = Spruch_des_Tages & "Leben, das ist das Allerseltenste in der Welt -" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die meisten Menschen existieren nur." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Oscar Wilde)" & vbCrLf

        Case 186

            Spruch_des_Tages = Spruch_des_Tages & "Tue soviel Gutes," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wie du kannst," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "und mache so wenig Gerede wie nur möglich darüber." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Charles Dickens)" & vbCrLf

        Case 187

            Spruch_des_Tages = Spruch_des_Tages & "Mütter lieben ihre Kinder mehr," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "als Väter es tun," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "weil sie sicher sein können," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass es ihre sind." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Aristoteles)" & vbCrLf
            
        Case 188

            Spruch_des_Tages = Spruch_des_Tages & "Als ich klein war, glaubte ich," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Geld sei das wichtigste im Leben." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Heute, da ich alt bin," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "weiß ich: Es stimmt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Oscar Wilde)" & vbCrLf
            
        Case 189

            Spruch_des_Tages = Spruch_des_Tages & "Erst wenn man genau weiß," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wie die Enkel ausgefallen sind," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "kann man beurteilen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "ob man seine Kinder gut erzogen hat." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Erich Maria Remarque)" & vbCrLf

        Case 190

            Spruch_des_Tages = Spruch_des_Tages & "Die meisten Menschen denken hauptsächlich über das nach," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "was die anderen Menschen über sie denken." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Sean Connery)" & vbCrLf



        Case 191

            Spruch_des_Tages = Spruch_des_Tages & "Es ist einfacher," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "kritisch zu sein als korrekt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Benjamin Disraeli)" & vbCrLf
            
        Case 192

            Spruch_des_Tages = Spruch_des_Tages & "Banken sind gefährlicher" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "als stehende Armeen" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Thomas Jefferson)" & vbCrLf
            
        Case 193

            Spruch_des_Tages = Spruch_des_Tages & "Humor ist keine Gabe des Geistes," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "er ist eine Gabe des Herzens." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Ludwig Börne)" & vbCrLf
            
        Case 194

            Spruch_des_Tages = Spruch_des_Tages & "Beim Flirten kommt es darauf an," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "eher die Notbremse zu ziehen als die Konsequenzen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Demi Moore)" & vbCrLf
            
        Case 195

            Spruch_des_Tages = Spruch_des_Tages & "Ein Optimist ist ein Mensch," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "der ein Dutzend Austern bestellt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "in der Hoffnung, sie mit der Perle," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die er darin findet, bezahlen zu können." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Theodor Fontane)" & vbCrLf
            
        Case 196

            Spruch_des_Tages = Spruch_des_Tages & "Machen Sie sich erst einmal unbeliebt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dann werden Sie auch ernst genommen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Konrad Adenauer)" & vbCrLf
            
        Case 197

            Spruch_des_Tages = Spruch_des_Tages & "Hast du keine Feinde," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dann hast du keinen Charakter." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Paul Newman)" & vbCrLf
            
        Case 198

            Spruch_des_Tages = Spruch_des_Tages & "Auch eine Enttäuschung," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wenn sie nur gründlich und endgültig ist," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "bedeutet einen Schritt vorwärts." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Max Planck)" & vbCrLf
            
        Case 199

            Spruch_des_Tages = Spruch_des_Tages & "Vertrauen wird dadurch erschöpft," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass es in Anspruch genommen wird." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Bertolt Brecht)" & vbCrLf
            
        Case 200

            Spruch_des_Tages = Spruch_des_Tages & "Ein Geiziger kann nichts Nützliches tun," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "als wenn er stirbt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Martin Luther)" & vbCrLf
            
        Case 201

            Spruch_des_Tages = Spruch_des_Tages & "Solange man selbst redet," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "erfährt man nichts." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Marie von Ebner-Eschenbach)" & vbCrLf
        
        Case 202

            Spruch_des_Tages = Spruch_des_Tages & "Ich möchte wie Ghandi sein" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "und wie Martin Luther King" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "und John Lennon." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Aber ich möchte am Leben bleiben." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Madonna)" & vbCrLf
            
        Case 203

            Spruch_des_Tages = Spruch_des_Tages & "Zeitverschwendung ist die leichteste " & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "aller Verschwendungen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Henry Ford)" & vbCrLf
        
        Case 204

            Spruch_des_Tages = Spruch_des_Tages & "Um eine Frau zu verführen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "muss man ihr nur einreden," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass ihr Ehemann sie nicht versteht." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Giacomo Casanova)" & vbCrLf
            
        Case 205

            Spruch_des_Tages = Spruch_des_Tages & "Erfolge muss man langsam löffeln," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "sonst verschluckt man sich an ihnen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Erika Pluhar)" & vbCrLf
        
        Case 206

            Spruch_des_Tages = Spruch_des_Tages & "Alles ist Kampf, Ringen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Nur der verdient die Liebe und das Leben," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "der täglich sie erobern muss." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Johann Wolfgang von Goethe)" & vbCrLf
            
        Case 207

            Spruch_des_Tages = Spruch_des_Tages & "Alles was du sagst, sollte wahr sein." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Aber nicht alles was wahr ist," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "solltest du auch sagen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Voltaire)" & vbCrLf
        
        Case 208

            Spruch_des_Tages = Spruch_des_Tages & "Die Welt ist ein Buch." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Wer nie reist," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "sieht nur eine Seite davon." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Augustinus Aurelius)" & vbCrLf
            
        Case 209

            Spruch_des_Tages = Spruch_des_Tages & "Frauen möchten in der Liebe Romane erleben," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Männer Kurzgeschichten." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Daphne du Maurier)" & vbCrLf
        
        Case 210

            Spruch_des_Tages = Spruch_des_Tages & "Es ist so leicht, andere," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "und so schwierig," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "sich selbst zu belehren." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Oscar Wilde)" & vbCrLf

        Case 211

            Spruch_des_Tages = Spruch_des_Tages & "Ehemänner sind vor allem dann gute Liebhaber," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wenn sie ihre Frauen betrügen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Marilyn Monroe)" & vbCrLf
            
        Case 212

            Spruch_des_Tages = Spruch_des_Tages & "Der Ziellose erleidet sein Schicksal" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "- der Zielbewusste gestaltet es." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Immanuel Kant)" & vbCrLf
            
        Case 213

            Spruch_des_Tages = Spruch_des_Tages & "Das Flüstern einer schönen Frau" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "hört man weiter als den lautesten Ruf der Pflicht." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Pablo Picasso)" & vbCrLf

        Case 214

            Spruch_des_Tages = Spruch_des_Tages & "Wenn du den Wert des Geldes kennenlernen willst," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "versuche, dir welches zu leihen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Benjamin Franklin)" & vbCrLf
            
        Case 215

            Spruch_des_Tages = Spruch_des_Tages & "Zwei Dinge sind unendlich," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "das Universum und die menschliche Dummheit," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "aber bei dem Universum" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "bin ich mir noch nicht ganz sicher." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Albert Einstein)" & vbCrLf
            
        Case 216

            Spruch_des_Tages = Spruch_des_Tages & "Liebst du das Leben?" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Dann vergeude keine Zeit," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "denn daraus besteht das Leben." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Benjamin Franklin)" & vbCrLf

        Case 217

            Spruch_des_Tages = Spruch_des_Tages & "Was andere uns zutrauen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "ist meist bezeichnender für sie als für uns." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Marie von Ebner-Eschenbach)" & vbCrLf
            
        Case 218

            Spruch_des_Tages = Spruch_des_Tages & "Gerne der Zeiten gedenk' ich," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "da alle Glieder gelenkig - bis auf eins." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Doch die Zeiten sind vorüber," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "steif geworden alle Glieder - bis auf eins." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Johann Wolfgang von Goethe)" & vbCrLf
            
        Case 219

            Spruch_des_Tages = Spruch_des_Tages & "Nur wer sein Alter verleugnet," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "fühlt sich wirklich alt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Lilli Palmer)" & vbCrLf

        Case 220

            Spruch_des_Tages = Spruch_des_Tages & "Ich bin besonders glücklich," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wenn das Glück unvollkommen ist." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Vollkommenheit hat keinen Charakter." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Peter Ustinov)" & vbCrLf
            
        Case 221

            Spruch_des_Tages = Spruch_des_Tages & "Erotik und Intelligenz müssen" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "nicht unbedingt Feinde sein." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Hildegard Knef)" & vbCrLf
            
        Case 222

            Spruch_des_Tages = Spruch_des_Tages & "Es ist besser, für das," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "was man ist, gehasst," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "als für das, was man nicht ist," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "geliebt zu werden." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(André Gide)" & vbCrLf

        Case 223

            Spruch_des_Tages = Spruch_des_Tages & "Je mehr man liebt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "um so tätiger wird man sein." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Vincent van Gogh)" & vbCrLf
            
        Case 224

            Spruch_des_Tages = Spruch_des_Tages & "Demokratie ist die Wahl" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "durch die beschränkte Mehrheit" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "anstelle der Ernennung" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "durch die bestechliche Minderheit." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(George Bernard Shaw)" & vbCrLf
            
        Case 225

            Spruch_des_Tages = Spruch_des_Tages & "Zum Mitleiden gab die Natur" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "vielen ein Talent," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "zur Mitfreude nur wenigen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Friedrich Hebbel)" & vbCrLf

        Case 226

            Spruch_des_Tages = Spruch_des_Tages & "Der kostbarste Besitz der Frau" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "ist die Phantasie des Mannes." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Beate Uhse)" & vbCrLf
            
        Case 227

            Spruch_des_Tages = Spruch_des_Tages & "Wenn Männer mein Dekolletee loben," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "freue ich mich." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Denn sonst werde ich zu sehr auf meine inneren Werte reduziert!" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Barbara Schöneberger)" & vbCrLf
            
        Case 228

            Spruch_des_Tages = Spruch_des_Tages & "Allem kann ich widerstehen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "nur der Versuchung nicht." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Oscar Wilde)" & vbCrLf
            
        Case 229

            Spruch_des_Tages = Spruch_des_Tages & "Ein Kluger bemerkt alles," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "ein Dummer macht über alles seine Bemerkungen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Heinrich Heine)" & vbCrLf

        Case 230

            Spruch_des_Tages = Spruch_des_Tages & "Das Geheimnis der Macht besteht darin," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "zu wissen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass andere noch feiger sind als wir." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Ludwig Börne)" & vbCrLf

        Case 231

            Spruch_des_Tages = Spruch_des_Tages & "Zynismus: ein Ding zu betrachten," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wie es wirklich ist," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "und nicht, wie es sein sollte." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Oscar Wilde)" & vbCrLf

        Case 232

            Spruch_des_Tages = Spruch_des_Tages & "Mit Kummer kann man allein fertig werden," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "aber um sich aus vollem Herzen freuen zu können," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "muss man die Freude teilen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Mark Twain)" & vbCrLf

        Case 233

            Spruch_des_Tages = Spruch_des_Tages & "Nicht ein Treuebruch ist das große Verbrechen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "sondern Gleichgültigkeit," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Bosheit und Intoleranz." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Lilli Palmer)" & vbCrLf

        Case 234

            Spruch_des_Tages = Spruch_des_Tages & "Was hilft aller Sonnenaufgang," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wenn wir nicht aufstehen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Georg Christoph Lichtenberg)" & vbCrLf
            
        Case 235

            Spruch_des_Tages = Spruch_des_Tages & "Nur ein mittelmäßiger Mensch" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "ist immer in Hochform." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(William Somerset Maugham)" & vbCrLf

        Case 236

            Spruch_des_Tages = Spruch_des_Tages & "Wie kann man einen Menschen beweinen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "der gestorben ist?" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Diejenigen sind zu beklagen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die ihn geliebt und verloren haben." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Helmuth von Moltke)" & vbCrLf

        Case 237

            Spruch_des_Tages = Spruch_des_Tages & "Alle menschlichen Organe werden irgendwann müde," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "nur die Zunge nicht." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Konrad Adenauer)" & vbCrLf

        Case 238

            Spruch_des_Tages = Spruch_des_Tages & "Wenn Du liebst," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "was Du tust," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wirst Du nie wieder in Deinem Leben arbeiten." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Konfuzius)" & vbCrLf

        Case 239

            Spruch_des_Tages = Spruch_des_Tages & "Einen Gescheiten kann man überzeugen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "einen Dummen muß man überreden." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Curt Goetz)" & vbCrLf

        Case 240

            Spruch_des_Tages = Spruch_des_Tages & "Einen Vorsprung im Leben hat," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wer da anpackt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wo die anderen erst einmal reden." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(John F.Kennedy)" & vbCrLf

        Case 241

            Spruch_des_Tages = Spruch_des_Tages & "Menschen, die immer daran denken," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "was andere von ihnen halten," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wären sehr überrascht, wenn sie wüßten," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wie wenig die anderen über sie nachdenken." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Bertrand Russell)" & vbCrLf
            
        Case 242

            Spruch_des_Tages = Spruch_des_Tages & "Der Optimist erklärt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass wir in der besten aller Welten leben," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "und der Pessimist fürchtet," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass dies wahr ist." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(James Branch Cabell)" & vbCrLf
            
        Case 243

            Spruch_des_Tages = Spruch_des_Tages & "Wer aufhört, besser werden zu wollen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "hört auf, gut zu sein." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Marie von Ebner-Eschenbach)" & vbCrLf
            
        Case 244

            Spruch_des_Tages = Spruch_des_Tages & "Der Kummer, der nicht spricht," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "nagt leise an dem Herzen, bis es bricht." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(William Shakespeare)" & vbCrLf
            
        Case 245

            Spruch_des_Tages = Spruch_des_Tages & "Es gibt keine Grenzen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Weder für Gedanken, noch für Gefühle." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Es ist die Angst, die immer Grenzen setzt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Ingmar Bergman)" & vbCrLf
            
        Case 246

            Spruch_des_Tages = Spruch_des_Tages & "Statt zu klagen, dass wir nicht alles haben," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "was wir wollen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "sollten wir lieber dankbar sein," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass wir nicht alles bekommen, was wir verdienen" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Dieter Hildebrandt)" & vbCrLf
            
        Case 247

            Spruch_des_Tages = Spruch_des_Tages & "Der Charakter ruht auf der Persönlichkeit," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "nicht auf den Talenten." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Johann Wolfgang von Goethe)" & vbCrLf
            
        Case 248

            Spruch_des_Tages = Spruch_des_Tages & "Wer all seine Ziele erreicht hat," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "hat sie sich als zu niedrig ausgewählt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Herbert von Karajan)" & vbCrLf



        Case 249

            Spruch_des_Tages = Spruch_des_Tages & "Kein Problem wird gelöst," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wenn wir träge darauf warten," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass Gott sich darum kümmert." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Martin Luther King)" & vbCrLf
            
        Case 250

            Spruch_des_Tages = Spruch_des_Tages & "Menschen zu finden," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die mit uns fühlen und empfinden," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "ist wohl das schönste Glück auf Erden." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Carl Spitteler)" & vbCrLf

        Case 251

            Spruch_des_Tages = Spruch_des_Tages & "Nur wer seinen eigenen Weg geht," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "kann von niemandem überholt werden." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Marlon Brando)" & vbCrLf
            
        Case 252

            Spruch_des_Tages = Spruch_des_Tages & "Meist belehrt erst der Verlust" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "über den Wert der Dinge." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Arthur Schopenhauer)" & vbCrLf
            
        Case 253

            Spruch_des_Tages = Spruch_des_Tages & "Ein Mensch," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "der für nichts zu sterben gewillt ist," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "verdient nicht zu leben." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Martin Luther King)" & vbCrLf
            
        Case 254

            Spruch_des_Tages = Spruch_des_Tages & "Es ist nicht schwer, Menschen zu finden," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die mit 60 Jahren zehnmal so reich sind," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "als sie es mit 20 waren." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Aber nicht einer von ihnen behauptet, er sei zehnmal so glücklich." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(George Bernard Shaw)" & vbCrLf
            
        Case 255

            Spruch_des_Tages = Spruch_des_Tages & "Alle Revolutionen haben bisher nur eines bewiesen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "nämlich, dass sich vieles ändern lässt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "bloß nicht die Menschen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Karl Marx)" & vbCrLf
            
        Case 256

            Spruch_des_Tages = Spruch_des_Tages & "Jeder möchte die Welt verbessern" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "und jeder könnte es auch," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wenn er nur bei sich selber anfangen wollte." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Karl Heinrich Waggerl)" & vbCrLf

        Case 257

            Spruch_des_Tages = Spruch_des_Tages & "Gegen Angriffe kann man sich wehren," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "gegen Lob ist man machtlos." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Sigmund Freud)" & vbCrLf

        Case 258

            Spruch_des_Tages = Spruch_des_Tages & "Wer sich zu groß fühlt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "um kleine Aufgaben zu erfüllen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "ist zu klein, um mit großen Aufgaben betraut zu werden." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Jacques Tati)" & vbCrLf

        Case 259

            Spruch_des_Tages = Spruch_des_Tages & "Was wir wissen, ist ein Tropfen;" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "was wir nicht wissen, ein Ozean." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Isaac Newton)" & vbCrLf

        Case 260

            Spruch_des_Tages = Spruch_des_Tages & "Man merkt nie, was schon getan wurde," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "man sieht immer nur, was noch zu tun bleibt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Marie Curie)" & vbCrLf

        Case 261

            Spruch_des_Tages = Spruch_des_Tages & "Wir gehen mit dieser Welt um," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "als hätten wir noch eine zweite im Kofferraum." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Jane Fonda)" & vbCrLf

        Case 262

            Spruch_des_Tages = Spruch_des_Tages & "Erfahrung ist nicht das," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "was einem zustößt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Erfahrung ist das," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "was man aus dem macht, was einem zustößt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Aldous Huxley)" & vbCrLf

        Case 263

            Spruch_des_Tages = Spruch_des_Tages & "Mensch: das einzige Lebewesen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "das erröten kann." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Es ist aber auch das einzige was Grund dazu hat." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Mark Twain)" & vbCrLf

        Case 264

            Spruch_des_Tages = Spruch_des_Tages & "Liebe mich dann," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wenn ich es am wenigsten verdient habe," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "denn dann brauche ich es am meisten." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Anonym)" & vbCrLf

        Case 265

            Spruch_des_Tages = Spruch_des_Tages & "Wenn die meisten sich schon armseliger Kleider und Möbel schämen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wieviel mehr sollten wir uns da erst" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "armseliger Ideen und Weltanschauungen schämen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Albert Einstein)" & vbCrLf

        Case 266

            Spruch_des_Tages = Spruch_des_Tages & "Wenn alle Menschen nur dann redeten," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wenn sie etwas zu sagen haben," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "würden sie bald den Gebrauch der Sprache verlieren." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(William Shakespeare)" & vbCrLf

        Case 267

            Spruch_des_Tages = Spruch_des_Tages & "Es ist traurig, eine Ausnahme zu sein." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Aber noch trauriger ist es, keine zu sein." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Peter Altenberg)" & vbCrLf

        Case 268

            Spruch_des_Tages = Spruch_des_Tages & "Nichts wird langsamer vergessen" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "als eine Beleidigung" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "und nichts eher als eine Wohltat." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Martin Luther)" & vbCrLf
            
        Case 269

            Spruch_des_Tages = Spruch_des_Tages & "Das Vergleichen ist das Ende des Glücks" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "und der Anfang der Unzufriedenheit." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Søren Kierkegaard)" & vbCrLf
    
        Case 270
            Spruch_des_Tages = Spruch_des_Tages & "Das Alter ist nicht trübe," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "weil darin unsere Freuden," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "sondern weil unsere Hoffnungen aufhören." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Jean Paul)"
            
        Case 271
            Spruch_des_Tages = Spruch_des_Tages & "Nur wer seine Rechnungen nicht bezahlt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "darf hoffen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "im Gedächtnis der Kaufleute weiterzuleben." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Oscar Wilde)" & vbCrLf
             
        Case 272
            Spruch_des_Tages = Spruch_des_Tages & "Seine eigene Dummheit zu erkennen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "mag schmerzhaft sein," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "keinesfalls aber eine Dummheit." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Oliver Hassencamp)" & vbCrLf
             
        Case 273
            Spruch_des_Tages = Spruch_des_Tages & "Wer immer mit dem Strom schwimmt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "erreicht niemals die Quelle." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Deutsches Sprichwort)" & vbCrLf
             
        Case 274
            Spruch_des_Tages = Spruch_des_Tages & "Zufriedenheit mit seiner Lage" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "ist der größte und sicherste Reichtum." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Marcus Tullius Cicero)" & vbCrLf
             
        Case 275
           
        
        
            Spruch_des_Tages = Spruch_des_Tages & "Moralische Entrüstung ist Neid" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "mit einem kleinen Heiligenschein." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Herbert George Wells)" & vbCrLf
             
        Case 276
            Spruch_des_Tages = Spruch_des_Tages & "Brüllt ein Mann, ist er dynamisch." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Brüllt eine Frau, ist sie hysterisch." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Hildegard Knef)" & vbCrLf
             
         Case 277
            Spruch_des_Tages = Spruch_des_Tages & "Was wir wissen, ist ein Tropfen;" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "was wir nicht wissen, ist ein Ozean." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Sir Isaac Newton)" & vbCrLf
             
         Case 278
            Spruch_des_Tages = Spruch_des_Tages & "Den Reiz des Verbotenen kann man nur kosten," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wenn man es sofort tut." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Morgen ist es vielleicht schon erlaubt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Jean Genet)" & vbCrLf
             
         Case 279
            Spruch_des_Tages = Spruch_des_Tages & "Wenn man die Inschriften auf den Friedhöfen liest," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "fragt man sich unwillkürlich," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wo denn eigentlich die Schurken begraben liegen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Peter Sellers)" & vbCrLf
             
        Case 280
            Spruch_des_Tages = Spruch_des_Tages & "Lesen ist für den Geist das," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "was Gymnastik für den Körper ist." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Joseph Addison)" & vbCrLf
             
        Case 281
            Spruch_des_Tages = Spruch_des_Tages & "Die gefährlichsten Unwahrheiten" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "sind Wahrheiten, mäßig entstellt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Georg Christoph Lichtenberg)" & vbCrLf
             
        Case 282
            Spruch_des_Tages = Spruch_des_Tages & "Man entdeckt keine neuen Erdteile," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "ohne den Mut zu haben," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "alte Küsten aus den Augen zu verlieren." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(André Gide)" & vbCrLf
             
         Case 283
            Spruch_des_Tages = Spruch_des_Tages & "Wer die Ursache nicht kennt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "nennt die Wirkung Zufall." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Werner Mitsch)" & vbCrLf
             
         Case 284
            Spruch_des_Tages = Spruch_des_Tages & "Die kürzesten Wörter," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "nämlich ja und nein," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "erfordern das meiste Nachdenken." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Pythagoras)" & vbCrLf
             
       Case 285
            Spruch_des_Tages = Spruch_des_Tages & "Erfahrungen sind Maßarbeit." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Sie passen nur dem, der sie macht." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Carlo Levi)" & vbCrLf
             
        Case 286
            Spruch_des_Tages = Spruch_des_Tages & "Unkraut ist die Opposition der Natur" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "gegen die Regierung der Gärtner." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Oskar Kokoschka)" & vbCrLf
             
        Case 287
            Spruch_des_Tages = Spruch_des_Tages & "Die Ehe ist eine lange Mahlzeit," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die mit dem Dessert beginnt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Henri de Toulouse-Lautrec)" & vbCrLf
             
        Case 288
            Spruch_des_Tages = Spruch_des_Tages & "Das Rezept für Gelassenheit ist ganz einfach:" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Man darf sich nicht über Dinge aufregen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die nicht zu ändern sind." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Helen Vita)" & vbCrLf
             
        Case 289
            Spruch_des_Tages = Spruch_des_Tages & "Alle Menschen sind klug:" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die einen vorher," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die anderen hinterher." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Chinesisches Sprichwort)" & vbCrLf
             
         Case 290
            Spruch_des_Tages = Spruch_des_Tages & "Jeder will alt werden," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "aber keiner will es sein." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Martin Held)" & vbCrLf
             
         Case 291
            Spruch_des_Tages = Spruch_des_Tages & "Gesegnet seien jene," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die nichts zu sagen haben" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "und den Mund halten." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Oscar Wilde)" & vbCrLf
             
         Case 292
            Spruch_des_Tages = Spruch_des_Tages & "Enten legen ihre Eier in aller Stille." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Hühner gackern dabei wie verrückt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Was ist die Folge?" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Alle Welt isst Hühnereier." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Henry Ford)" & vbCrLf
             
        Case 293
            Spruch_des_Tages = Spruch_des_Tages & "Jeder Mensch macht Fehler." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Das Kunststück liegt darin," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "sie zu machen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wenn keiner zuschaut." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Sir Peter Ustinov)" & vbCrLf
             
        Case 294
            Spruch_des_Tages = Spruch_des_Tages & "Die Klage über die Stärke des Wettbewerbs" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "ist in Wirklichkeit meist nur eine Klage" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "über den Mangel an eigenen Einfällen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Walther Rathenau)" & vbCrLf
             
         Case 295
            Spruch_des_Tages = Spruch_des_Tages & "Man kann niemanden überholen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wenn man in seine Fußstapfen tritt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Francois Truffaut)" & vbCrLf
             
         Case 296
            Spruch_des_Tages = Spruch_des_Tages & "Wer den Mund hält, weil er Unrecht hat," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "ist ein Weiser." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Wer den Mund hält, obwohl er Recht hat," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "ist verheiratet oder Pfeifenraucher." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(George Bernard Shaw)" & vbCrLf
             
         Case 297
            Spruch_des_Tages = Spruch_des_Tages & "Das größte Übel der heutigen Jugend besteht darin," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass man nicht mehr dazugehört." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Salvador Dali)" & vbCrLf
             
         Case 298
            Spruch_des_Tages = Spruch_des_Tages & "Wenn die Pflicht ruft," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "gibt es viele Schwerhörige." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Gustav Knuth)" & vbCrLf
             
       Case 299
            Spruch_des_Tages = Spruch_des_Tages & "Es gibt nur eine Ausflucht von der Arbeit:" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Andere für sich arbeiten lassen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Immanuel Kant)" & vbCrLf
             
        Case 300
            Spruch_des_Tages = Spruch_des_Tages & "Die meisten Menschen wären glücklich," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wenn sie sich das Leben leisten könnten," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "das sie sich leisten." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Danny Kaye)" & vbCrLf
             
        Case 301
            Spruch_des_Tages = Spruch_des_Tages & "Ein Tag, an dem man nicht lacht," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "ist ein verlorener Tag." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Charlie Chaplin)" & vbCrLf
             
        Case 302
            Spruch_des_Tages = Spruch_des_Tages & "Zwei Dinge sind unendlich:" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Das Universum und die menschliche Dummheit." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Aber beim Universum bin ich mir nicht ganz sicher." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Albert Einstein)" & vbCrLf
             
        Case 303
            Spruch_des_Tages = Spruch_des_Tages & "Auch wenn 50 Millionen Menschen" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "etwas Dummes sagen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "bleibt es trotzdem eine Dummheit." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Anatole France)" & vbCrLf
             
         Case 304
            Spruch_des_Tages = Spruch_des_Tages & "Die Ehe ist der Versuch," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "zu zweit mit den Problemen fertig zu werden," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die man allein nie gehabt hätte." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Woody Allen)" & vbCrLf
             
       Case 305
            Spruch_des_Tages = Spruch_des_Tages & "Die meisten Menschen verwenden" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "mehr Zeit und Kraft darauf," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "über Probleme zu diskutieren" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "statt sie anzupacken." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Henry Ford)" & vbCrLf
             
        Case 306
            Spruch_des_Tages = Spruch_des_Tages & "Um klar zu sehen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "genügt oft ein Wechsel der Blickrichtung." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Antoine de Saint-Exupéry)" & vbCrLf
             
        Case 307
            Spruch_des_Tages = Spruch_des_Tages & "Als du auf die Welt kamst," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "hast du geweint," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "um dich herum freuten sich alle." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Lebe so, dass wenn du die Welt verlässt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "alle weinen und du allein lächelst." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Chinesisches Sprichwort)" & vbCrLf
             
         Case 308
            Spruch_des_Tages = Spruch_des_Tages & "Es gibt keine Passagiere" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "auf dem Raumschiff Erde," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "jeder gehört zur Besatzung." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Marshall McLuhan)" & vbCrLf
             
        Case 309
            Spruch_des_Tages = Spruch_des_Tages & "Wer den Menschen die Hölle auf Erden" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "bereiten will," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "braucht ihnen nur alles zu erlauben." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Graham Greene)" & vbCrLf
             
        Case 310
            Spruch_des_Tages = Spruch_des_Tages & "Ein neuer Gedanke wird zuerst verlacht," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dann bekämpft," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "bis er nach längerer Zeit" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "als selbstverständlich gilt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Arthur Schopenhauer)" & vbCrLf
             
        Case 311
            Spruch_des_Tages = Spruch_des_Tages & "Obwohl sie nicht einmal" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "hundert Jahre alt werden," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "bereiten sich die Menschen" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Sorgen für tausend Jahre." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Chinesisches Sprichwort)" & vbCrLf
             
        Case 312
            Spruch_des_Tages = Spruch_des_Tages & "Es gehört zu den alltäglichen Täuschungen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die Stunden der Vergangenheit und Zukunft" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "reizender zu finden als die Gegenwart." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Heinrich Zschokke)" & vbCrLf
             
        Case 313
            Spruch_des_Tages = Spruch_des_Tages & "Das beste Mittel, jeden Tag gut zu beginnen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "ist, beim Erwachen daran zu denken," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "ob man nicht wenigstens einem Menschen" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "an diesem Tag eine Freude machen könne." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Friedrich Nietzsche)" & vbCrLf
             
        Case 314
            Spruch_des_Tages = Spruch_des_Tages & "Wenn man einen Menschen richtig beurteilen will," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "so frage man sich immer:" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Willst du den zum Vorgesetzten haben?" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Kurt Tucholsky)" & vbCrLf
             
        Case 315
            Spruch_des_Tages = Spruch_des_Tages & "Wir leben in einem gefährlichen Zeitalter." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Der Mensch beherrscht die Natur," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "bevor er gelernt hat," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "sich selber zu beherrschen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Albert Schweitzer)" & vbCrLf
             
       Case 316
            Spruch_des_Tages = Spruch_des_Tages & "Je schlechter der Redner," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "desto länger seine Rede." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Japanisches Sprichwort)" & vbCrLf
             
        Case 317
            Spruch_des_Tages = Spruch_des_Tages & "Willst du den Charakter" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "eines Menschen erkennen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "so gib ihm Macht!" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Abraham Lincoln)" & vbCrLf
             
         Case 318
            Spruch_des_Tages = Spruch_des_Tages & "Der Hauptfehler des Menschen bleibt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass er so viele kleine hat." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Jean Paul)" & vbCrLf
             
         Case 319
            Spruch_des_Tages = Spruch_des_Tages & "Mit Adleraugen sehen wir die Fehler anderer," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "mit Maulwurfsaugen unsere eigenen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Franz von Sales)" & vbCrLf
             
        Case 320
            Spruch_des_Tages = Spruch_des_Tages & "Wenn du wissen willst," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "was dein Nachbar von dir denkt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "so fange Streit mit ihm an." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Afrikanisches Sprichwort)" & vbCrLf
             
         Case 321
            Spruch_des_Tages = Spruch_des_Tages & "Es ist von großem Vorteil," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die Fehler, aus denen man lernen kann," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "recht frühzeitig zu machen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Winston Churchill)" & vbCrLf
             
          Case 322
            Spruch_des_Tages = Spruch_des_Tages & "Ich habe immer gefunden," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass die Türen, durch die ich gehen soll," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "sich mir von selbst öffnen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Gewaltsam durchzudringen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "ist mir nie gut bekommen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Robert Wilhelm Bunsen)" & vbCrLf
             
        Case 323
            Spruch_des_Tages = Spruch_des_Tages & "Begangene Fehler können" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "nicht besser entschuldigt werden" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "als mit dem Geständnis," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass man sie als solche erkenne." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Pedro Calderón de la Barca)" & vbCrLf
             
         Case 324
            Spruch_des_Tages = Spruch_des_Tages & "Es gibt Dinge," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die man bereut, ehe man sie tut." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Und man tut sie doch." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Christian Friedrich Hebbel)" & vbCrLf
             
         Case 325
            Spruch_des_Tages = Spruch_des_Tages & "Ein Lügner muss ein gutes Gedächtnis haben." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Spanisches Sprichwort)" & vbCrLf
             
         Case 326
            Spruch_des_Tages = Spruch_des_Tages & "Jede Rohheit hat ihre Ursache in einer Schwäche." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Lucius Annaeus Seneca)" & vbCrLf
             
         Case 327
            Spruch_des_Tages = Spruch_des_Tages & "Jeder Fehler erscheint unglaublich dumm," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wenn and're ihn begehen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Georg Christoph Lichtenberg)" & vbCrLf
             
         Case 328
            Spruch_des_Tages = Spruch_des_Tages & "Unsere Prinzipien dauern gerade so lange," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "bis sie mit unseren Leidenschaften" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "oder Eitelkeiten in Konflikt kommen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "und ziehen dann jedes Mal den Kürzeren." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Theodor Fontane)" & vbCrLf
 
        Case 329
            Spruch_des_Tages = Spruch_des_Tages & "Die Kunst des Umgangs mit Menschen besteht darin," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "sich geltend zu machen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "ohne andere unerlaubt zurück zu drängen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Adolf Freiherr von Knigge)" & vbCrLf
        
        Case 330
            Spruch_des_Tages = Spruch_des_Tages & "Während die Weisen grübeln," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "erobern die Dummen die Festung." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Serbisches Sprichwort)" & vbCrLf
            
        Case 331
            Spruch_des_Tages = Spruch_des_Tages & "Wenn man lange genug wartet," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "kommt das schönste Wetter." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Japanisches Sprichwort)" & vbCrLf
            
       Case 332
            Spruch_des_Tages = Spruch_des_Tages & "Wer zu schwach ist, dir zu nutzen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "ist noch immer stark genug, dir zu schaden." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Deutsches Sprichwort)" & vbCrLf
            
       Case 333
            Spruch_des_Tages = Spruch_des_Tages & "Leidenschaften gleichen Blendlaternen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Sie werfen alles Licht nach einer Richtung," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "während alles andere rings im Dunkel bleibt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Peter Sirius)" & vbCrLf
       Case 334
            
            Spruch_des_Tages = Spruch_des_Tages & "Ein kurzer Augenblick der Seelenruhe ist besser" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "als alles, was du sonst erstreben magst." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Persisches Sprichwort)" & vbCrLf
       Case 335
            
            Spruch_des_Tages = Spruch_des_Tages & "Das waren noch glückliche Zeiten," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "da man nach dem Kalender lebte!" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Heute lebt man nach der Uhr." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Sacha Guitry)" & vbCrLf
            
        Case 336
            Spruch_des_Tages = Spruch_des_Tages & "Wir sind nicht nur verantwortlich für das, was wir tun," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "sondern auch für das, was wir nicht tun." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Jean Baptiste Molière)" & vbCrLf
            
        Case 337
            Spruch_des_Tages = Spruch_des_Tages & "Eine Lüge ist wie ein Schneeball:" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Je länger man ihn wälzt, desto größer wird er." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Martin Luther)" & vbCrLf
            
        Case 338
            Spruch_des_Tages = Spruch_des_Tages & "Zu dem, der warten kann," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "kommt alles mit der Zeit." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Französisches Sprichwort)" & vbCrLf
            
         Case 339
            Spruch_des_Tages = Spruch_des_Tages & "Schön ist eigentlich alles," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "was man mit Liebe betrachtet." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Christian Morgenstern)" & vbCrLf
            
         Case 340
            Spruch_des_Tages = Spruch_des_Tages & "Es ist ein grundsätzlicher Irrtum," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Heftigkeit und Starrheit Stärke zu nennen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Thomas Carlyle)" & vbCrLf
            
         Case 341
            Spruch_des_Tages = Spruch_des_Tages & "Vier Dinge sind es, die nicht zurück kommen:" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "das gesprochene Wort," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "der abgeschossene Pfeil," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "das vergangene Leben und" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "die versäumte Gelegenheit." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Unbekannter Verfasser)" & vbCrLf
            
         Case 342
            Spruch_des_Tages = Spruch_des_Tages & "Es genügt nicht, an den Fluss zu gehen" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "mit dem Wunsch, Fische zu fangen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Man muss auch ein Netz mitbringen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Koreanisches Sprichwort)" & vbCrLf
            
         Case 343
            Spruch_des_Tages = Spruch_des_Tages & "Auch eine Reise von 1000 Meilen" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "fängt mit dem ersten Schritt an." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Chinesisches Sprichwort)" & vbCrLf
            
         Case 344
            Spruch_des_Tages = Spruch_des_Tages & "Die Heirat ist die einzige lebenslängliche Verurteilung," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "bei der man auf Grund schlechter Führung" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "begnadigt werden kann." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Alfred Hitchcock)" & vbCrLf
            
         Case 345
            Spruch_des_Tages = Spruch_des_Tages & "Fordere viel von dir selbst" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "und erwarte wenig von den anderen!" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Konfuzius)" & vbCrLf
            
         Case 346
            Spruch_des_Tages = Spruch_des_Tages & "Man kann es auf zweierlei Weise zu etwas bringen:" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Durch eigenes Können" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "oder durch die Dummheit der anderen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Jean de La Bruyère)" & vbCrLf
            
        Case 347
            Spruch_des_Tages = Spruch_des_Tages & "Den Spiegel darfst du nicht schelten," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wenn er dir eine schiefe Fratze zeigt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Russisches Sprichwort)" & vbCrLf
            
        Case 348
            Spruch_des_Tages = Spruch_des_Tages & "Alles, was wir wirklich lieben, ist unersetzlich," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "und alles, wofür Ersatz nur denkbar ist," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "haben wir niemals wahrhaftig geliebt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Gustav Nieritz)" & vbCrLf
            
        Case 349
            Spruch_des_Tages = Spruch_des_Tages & "Der Flirt ist die Kunst," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "einer Frau in die Arme zu sinken," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "ohne ihr in die Hände zu fallen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Sacha Guitry)" & vbCrLf
            
        Case 350
            Spruch_des_Tages = Spruch_des_Tages & "Wie kahl und jämmerlich" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "würde manches Stück Erde aussehen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wenn kein Unkraut darauf wüchse." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Wilhelm Raabe)" & vbCrLf
            
         Case 351
            Spruch_des_Tages = Spruch_des_Tages & "Wer im Ruf steht, ein Frühaufsteher zu sein," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "kann getrost den ganzen Morgen im Bett bleiben." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Holländisches Sprichwort)" & vbCrLf
            
         Case 352
            Spruch_des_Tages = Spruch_des_Tages & "Die Welt ist wie ein Brei." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Zieht man den Löffel heraus, und wär's der größte," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "gleich klappt die Geschichte wieder zusammen," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "als ob gar nichts passiert wäre." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Wilhelm Busch)" & vbCrLf
            
        Case 353
            Spruch_des_Tages = Spruch_des_Tages & "Die Leute, die niemals Zeit haben," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "tun am wenigsten." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Georg Christoph Lichtenberg)" & vbCrLf
            
        Case 354
            Spruch_des_Tages = Spruch_des_Tages & "Es ist nicht wenig Zeit, was wir haben," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "sondern es ist viel, was wir nicht nützen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Lucius Annaeus Seneca)" & vbCrLf
         Case 355
            
            Spruch_des_Tages = Spruch_des_Tages & "Freundlich abschlagen ist besser" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "als unwillig geben." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Unbekannter Verfasser)" & vbCrLf
            
        Case 356
            Spruch_des_Tages = Spruch_des_Tages & "Wenn du den Hahn einsperrst," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "geht die Sonne doch auf." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Indisches Sprichwort)" & vbCrLf
            
        Case 357
            Spruch_des_Tages = Spruch_des_Tages & "Der Nachteil der Intelligenz besteht darin," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "dass man ständig gezwungen ist, dazuzulernen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(George Bernard Shaw)" & vbCrLf
            
       Case 358
            Spruch_des_Tages = Spruch_des_Tages & "Was mich nicht umbringt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "macht mich noch stärker." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Friedrich Nietzsche)" & vbCrLf
            
       Case 359
            Spruch_des_Tages = Spruch_des_Tages & "Der Reichtum gleicht dem Seewasser:" & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Je mehr man davon trinkt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "desto durstiger wird man." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Arthur Schopenhauer)" & vbCrLf
        Case 360
            
            Spruch_des_Tages = Spruch_des_Tages & "Der gute Ruf geht weit," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "aber der schlechte noch viel weiter." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Serbisches Sprichwort)" & vbCrLf
            
       Case 361
            
            Spruch_des_Tages = Spruch_des_Tages & "Tue Gutes: Dein Nachbar erfährt es nie." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "Tue Böses: Man weiß es auf hundert Meilen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Chinesisches Sprichwort)" & vbCrLf
            
       Case 362
            
            Spruch_des_Tages = Spruch_des_Tages & "Man wird des Guten und auch des Besten," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wenn es alltäglich zu sein beginnt," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "bald satt." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Gotthold Ephraim Lessing)" & vbCrLf
            
       Case 363
            
            Spruch_des_Tages = Spruch_des_Tages & "Wirf deinen Lendenschurz nicht weg," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wenn du ein neues Kleid bekommen hast." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Afrikanisches Sprichwort)" & vbCrLf
            
        Case 364
            
            Spruch_des_Tages = Spruch_des_Tages & "Dem Ersten gebührt der Ruhm," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "auch wenn die Nachfolger es besser gemacht haben." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Arabisches Sprichwort)" & vbCrLf
            
       Case 365
            
            Spruch_des_Tages = Spruch_des_Tages & "Wer immer nur wartet," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "bis ein anderer ihm zum Essen ruft," & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "wird oft nichts zu essen bekommen." & vbCrLf
            Spruch_des_Tages = Spruch_des_Tages & "(Rumänisches Sprichwort)" & vbCrLf
        
    End Select
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "Spruch_des_Tages"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function pfadaendernplusDatname(sTitle As String, sFilter As String, sOldpfad As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim sPfad   As String
    
    pfadaendernplusDatname = sOldpfad
    
    With frmWKL00.cdlopen
        .CancelError = True
        On Error GoTo err
        .InitDir = sOldpfad
        .DialogTitle = sTitle
        .Filter = sFilter
        .ShowSave
    
        sPfad = .FileName
    End With
    pfadaendernplusDatname = sPfad
err:
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "pfadaendernplusDatname"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function ermBestandfromZbestand(cArtNr As String, iFil As Integer) As Long
On Error GoTo LOKAL_ERROR

Dim sSQL As String
Dim rsrs As Recordset

ermBestandfromZbestand = 0

sSQL = "Select * from ZBESTAND where artnr = " & cArtNr
sSQL = sSQL & " and Filialnr = " & iFil

Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    If Not IsNull(rsrs!BESTAND) Then
        ermBestandfromZbestand = rsrs!BESTAND
    End If
End If
rsrs.Close

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermBestandfromZbestand"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub lWEinBESTLIN(lblanzeige As Label)
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim rsUW        As Recordset
    Dim cLinr       As String
    Dim lMaxDate    As Long

    Screen.MousePointer = 11
    
    sSQL = " Select * from BESTLIN where MAXDATE = clng(datevalue('00:00:00')) "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!linr) Then
                cLinr = rsrs!linr
                'datum lzugang
                lMaxDate = 0
                
                sSQL = "select max(adate) as DATN  from zugang "
                sSQL = sSQL & " Where linr = " & cLinr
                
                Set rsUW = gdBase.OpenRecordset(sSQL)
                If Not rsUW.EOF Then
                    If Not IsNull(rsUW!DATN) Then
                        lMaxDate = rsUW!DATN
                    End If
                End If
                rsUW.Close
                
                sSQL = "Update BESTLIN set MAXDATE = " & lMaxDate
                sSQL = sSQL & " where linr = " & cLinr
                gdBase.Execute sSQL, dbFailOnError
                
                'datum lzugang
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    anzeige "normal", "", lblanzeige
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "lWEinBESTLIN"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function Zeitungs_EAN_ZU_Artnr(sZeitungsean As String, sMWST As String) As Long
On Error GoTo LOKAL_ERROR

    Dim lLinr As Long

    Zeitungs_EAN_ZU_Artnr = 0

    lLinr = glZeitungsLinr ' ermLinrInZeitE
    
    If lLinr > 0 Then
        Zeitungs_EAN_ZU_Artnr = Val(ermartnrausLIBESNR(CStr(Val(Mid(sZeitungsean, 4, 5))), lLinr))
    Else
        Zeitungs_EAN_ZU_Artnr = 0
    End If

    If sMWST = "E" Then
        If Zeitungs_EAN_ZU_Artnr = 0 Then Zeitungs_EAN_ZU_Artnr = 666668
    ElseIf sMWST = "V" Then
        If Zeitungs_EAN_ZU_Artnr = 0 Then Zeitungs_EAN_ZU_Artnr = 666669
    End If
        

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "Zeitungs_EAN_ZU_Artnr"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function

Public Sub ErmittleBestVorschlag(lblanzeige As Label)
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim rsUW        As Recordset
    Dim cLinr       As String
    Dim i           As Integer
    Dim j           As Integer
    Dim lcount      As Long
    Dim lsumKB      As Long
    Dim lsumFART    As Long
    Dim lanzBe      As Long
    Dim iRet        As Long
    Dim dWertLEK    As Double
    Dim lMaxDate    As Long

    Screen.MousePointer = 11

    loeschNEW "BESTLIN", gdBase
    CreateTableT2 "BESTLIN", gdBase

    Dim lVon    As Long
    Dim lBis    As Long
    
    Dim lDiff1  As Long
    Dim lDiff2  As Long
    Dim lDif    As Long
    
    Dim iTage As Integer
    

    If ErmMBSTAND <= CLng(DateValue(Now) - 3) Then
    
        leseMBDetails
    
        Select Case MBDETAILMON
            Case 5 '9
                iTage = 272
            Case 4 '8
                iTage = 241
            Case 3 '7
                iTage = 211
            Case 2 '6
                iTage = 180
            Case 1 '5
                iTage = 150
            Case 0 '4
                iTage = 119
            Case Else
                iTage = 180
        End Select
    
        lVon = DateValue(Now) - iTage
        lBis = DateValue(Now)
        
        lDiff1 = lBis - lVon
        lDiff2 = MBDETAILBIS - MBDETAILVON
        
        If MBDETAILVON <= lBis And MBDETAILBIS <= lBis And MBDETAILVON >= lVon And MBDETAILBIS >= lVon Then
            lDif = lDiff1 - lDiff2
        ElseIf MBDETAILVON <= lBis And MBDETAILBIS > lBis And MBDETAILVON >= lVon Then
            lDif = MBDETAILVON - lVon
        ElseIf MBDETAILBIS <= lBis And MBDETAILVON < lVon And MBDETAILBIS >= lVon Then
            lDif = lBis - MBDETAILBIS
        ElseIf MBDETAILVON < lVon And MBDETAILBIS < lVon Then
            lDif = lDiff1
            
        ElseIf MBDETAILVON > lBis And MBDETAILBIS > lBis Then
            lDif = lDiff1
        End If
        
        MBrechnen1 MBDETAILBVO, CInt(lDif), 1, lVon, lBis, lblanzeige, MBDETAILVON, MBDETAILBIS
    
    End If
    
    
    
    
   
    
    sSQL = " Select * from Lisrt where kuerzel <> '' or not kuerzel is null "
'    sSQL = " Select * from Lisrt where linr = 312130"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lcount = rsrs.RecordCount
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!linr) Then
                cLinr = rsrs!linr
                anzeige "normal", cLinr & " (" & lcount & ")", lblanzeige
                lcount = lcount - 1
            
                loeschNEW "BESTART", gdBase
                CreateTableT2 "BESTART", gdBase
                
                sSQL = "Insert into BESTART select artnr "
                sSQL = sSQL & " ,Bestand "
                sSQL = sSQL & " ,MINBEST as MB "
                sSQL = sSQL & " from artikel where linr = " & cLinr
                sSQL = sSQL & " and RKZ = 'N' "
                sSQL = sSQL & " and Gefuehrt = 'J' "
                gdBase.Execute sSQL, dbFailOnError
                
'                sSQL = "Delete from BESTART where   "
'                sSQL = sSQL & " BESTART.artnr in (select SF" & srechnertab & ".artnr from SF" & srechnertab & ")"
'                sSQL = sSQL & " and Filiale = " & i
'                gdBase.Execute sSQL, dbFailOnError
                    
                
                
                j = 1
                anzeige "normal", lblanzeige & ".." & j, lblanzeige
                j = j + 1
                
'                sSQL = "Update BESTART "
'                sSQL = sSQL & " set bestart.bestand = 0"
'                sSQL = sSQL & " , bestart.MB = 0"
'                gdBase.Execute sSQL, dbFailOnError
                
''                sSQL = "Update BESTART inner join zbestand on BESTART.ARTNR = ZBESTAND.ARTNR "
''                sSQL = sSQL & " and BESTART.FILIALE = ZBESTAND.FILIALNR "
''                sSQL = sSQL & " set bestart.bestand = zbestand.bestand"
''                sSQL = sSQL & " , bestart.MB = zbestand.MINBEST"
''                gdBase.Execute sSQL, dbFailOnError
                
                'hier Bedarf erhöhen wegen Kundenbestellungen
                
                sSQL = "Update BESTART set KB = 0 "
                gdBase.Execute sSQL, dbFailOnError
                
                sSQL = "Update BESTART inner join KUNDBEST on BESTART.ARTNR = KUNDBEST.ARTNR "
'                sSQL = sSQL & " and BESTART.FILIALE = KUNDBEST.FILIALE "
                sSQL = sSQL & " set bestart.KB = KUNDBEST.BESTELLTMENGE"
                sSQL = sSQL & " where  KUNDBEST.StatusARTIKEL = 'INBESTELLUNG' "
                gdBase.Execute sSQL, dbFailOnError
                
'                anzeige "normal", lblanzeige & ".." & j, lblanzeige
'                j = j + 1
                
                'hier gesperrte außer Kundenbestellungen raus
                
'                sSQL = "Delete from BESTART where   "
'                sSQL = sSQL & " BESTART.artnr in (select SF" & srechnertab & ".artnr from SF" & srechnertab & ")"
'                sSQL = sSQL & " and  BESTART.KB = 0 "
'                gdBase.Execute sSQL, dbFailOnError
                
                anzeige "normal", lblanzeige & ".." & j, lblanzeige
                j = j + 1
                
                'hier die unterwegs sind ermitteln
                sSQL = "Update BESTART set UW = 0 "
                gdBase.Execute sSQL, dbFailOnError
                
                
                
                anzeige "normal", lblanzeige & ".." & j, lblanzeige
                j = j + 1
                
               
            
'                sSQL = "Select * from BESTART "
'                Set rsUW = gdBase.OpenRecordset(sSQL)
'                If Not rsUW.EOF Then
'                    rsUW.MoveFirst
'                    Do While Not rsUW.EOF
'
'                    If Not IsNull(rsUW!artnr) Then
'
'                        rsUW.Edit
'                        rsUW!UW = LeseUnterwegs(CLng(rsUW!artnr), CLng(i))
'                        rsUW.Update
'
'                    End If
'
'                    rsUW.MoveNext
'                    Loop
'                End If
'                rsUW.Close



                
                
                
                anzeige "normal", lblanzeige & ".." & j, lblanzeige
                j = j + 1
                
                sSQL = "Update BESTART "
                sSQL = sSQL & " set bestand = bestand + uw "
                gdBase.Execute sSQL, dbFailOnError
                
                sSQL = "Update BESTART "
                sSQL = sSQL & " set MB = MB + KB "
                gdBase.Execute sSQL, dbFailOnError
    
                sSQL = "Update BESTART "
                sSQL = sSQL & " set bestand = 0"
                sSQL = sSQL & " where bestand < 0"
                gdBase.Execute sSQL, dbFailOnError
                
                sSQL = "Update BESTART "
                sSQL = sSQL & " set bedarf =  MB - bestand "
                gdBase.Execute sSQL, dbFailOnError
                
                sSQL = "Update BESTART "
                sSQL = sSQL & " set Uberbest =  bestand - MB"
                gdBase.Execute sSQL, dbFailOnError
                
                sSQL = "Update BESTART "
                sSQL = sSQL & " set Uberbest = 0 where UBERBEST < 0 "
                gdBase.Execute sSQL, dbFailOnError
                
                anzeige "normal", lblanzeige & ".." & j, lblanzeige
                j = j + 1
                
                loeschNEW "BESTART2", gdBase
                CreateTableT2 "BESTART2", gdBase
                
                sSQL = "Insert into BESTART2 select artnr"
                sSQL = sSQL & " , sum(Uberbest) as Uberbest1 "
                sSQL = sSQL & " from BESTART "
                sSQL = sSQL & " where Filiale <> 1 "
                sSQL = sSQL & " group by artnr "
                gdBase.Execute sSQL, dbFailOnError
                
                sSQL = "Delete from BESTART "
                sSQL = sSQL & " where bedarf <= 0"
'                sSQL = sSQL & " and  filiale  <> " & gbyteFilnr ' das kam am 070607 dazu
                gdBase.Execute sSQL, dbFailOnError
                
                'Teil2 gesamtbedarf pro artikel und bestand in 1
                
                loeschNEW "BESTART1", gdBase
                CreateTableT2 "BESTART1", gdBase
                
                sSQL = "Insert into BESTART1 select artnr"
                sSQL = sSQL & " , sum(bedarf) as bedarf1 "
                sSQL = sSQL & " , bestand "
                sSQL = sSQL & " , 0 as KONDI "
                sSQL = sSQL & " from BESTART "
                sSQL = sSQL & " group by artnr,bestand "
                gdBase.Execute sSQL, dbFailOnError
                
                anzeige "normal", lblanzeige & ".." & j, lblanzeige
                j = j + 1
                
                
                Set rsUW = gdBase.OpenRecordset("BESTART1")
                If Not rsUW.EOF Then
                    rsUW.MoveFirst
                    Do While Not rsUW.EOF
            
                    If Not IsNull(rsUW!artnr) Then
                        rsUW.Edit
                        rsUW!kondi = 0 'KondiPLUS(80, rsUW!ARTNR, CInt(rsUW!Bedarf1))
                        rsUW.Update
                    End If
            
                    rsUW.MoveNext
                    Loop
                End If
                rsUW.Close
                
                anzeige "normal", lblanzeige & ".." & j, lblanzeige
                j = j + 1
                
                sSQL = "Update BESTART1 "
                sSQL = sSQL & " set bedarf1 =  KONDI "
                sSQL = sSQL & " where  KONDI > 0 "
                gdBase.Execute sSQL, dbFailOnError
                
                
                
                Set rsUW = gdBase.OpenRecordset("BESTART1")
                If Not rsUW.EOF Then
                    rsUW.MoveFirst
                    Do While Not rsUW.EOF
            
                    If Not IsNull(rsUW!artnr) Then
                        rsUW.Edit
                        rsUW!LEK = ermLEKPR(rsUW!artnr, CLng(cLinr))
                        rsUW.Update
                    End If
            
                    rsUW.MoveNext
                    Loop
                End If
                rsUW.Close
                
                
                
                Set rsUW = gdBase.OpenRecordset("BESTART1")
                If Not rsUW.EOF Then
                    rsUW.MoveFirst
                    Do While Not rsUW.EOF
            
                    If Not IsNull(rsUW!artnr) Then
                        rsUW.Edit
                        rsUW!inBe = erminBestell(rsUW!artnr)
                        rsUW.Update
                    End If
            
                    rsUW.MoveNext
                    Loop
                End If
                rsUW.Close
                
                anzeige "normal", lblanzeige & ".." & j, lblanzeige
                j = j + 1
                
                'vorher in bestellung abziehen
                
                sSQL = "Update BESTART1 set bedarf1 = bedarf1 - INBE "
                gdBase.Execute sSQL, dbFailOnError
                
                sSQL = "Update BESTART1 set bedarf1 = 0 Where bedarf1 < 0 "
                gdBase.Execute sSQL, dbFailOnError
                
                sSQL = "Update BESTART1 set LEKWERT =  LEK * bedarf1 "
                gdBase.Execute sSQL, dbFailOnError
                
                sSQL = "Update BESTART1 set INBELEKWERT =  LEK * INBE "
                gdBase.Execute sSQL, dbFailOnError
                
                
                
                'Wenn alles fertig - dann in die bestlin
                sSQL = "Insert into BESTLIN select " & cLinr & " as linr "
                sSQL = sSQL & " , sum(bedarf1) as ANZART "
                sSQL = sSQL & " , sum(LEKWERT) as VORLEK "
'                sSQL = sSQL & " , sum(INBELEKWERT) as INBESTLEK "
                sSQL = sSQL & " from BESTART1 "
                gdBase.Execute sSQL, dbFailOnError
                
                anzeige "normal", lblanzeige & ".." & j, lblanzeige
                j = j + 1
                
                'Anzahl Aufträge
                lanzBe = 0
                
                loeschNEW "COUNTBE", gdBase
                sSQL = "select distinct(DATEINAME) as DATN into COUNTBE from bestrest "
                sSQL = sSQL & " Where linr = " & cLinr
                gdBase.Execute sSQL, dbFailOnError
                
                sSQL = "select count(DATN) as MAXI from COUNTBE "
                
                Set rsUW = gdBase.OpenRecordset(sSQL)
                If Not rsUW.EOF Then
                    If Not IsNull(rsUW!maxi) Then
                        lanzBe = rsUW!maxi
                    End If
                End If
                rsUW.Close
                
                sSQL = "Update BESTLIN set anzbe = " & lanzBe
                sSQL = sSQL & " where linr = " & cLinr
                gdBase.Execute sSQL, dbFailOnError
                
                'LEK WERT Aufträge
                dWertLEK = 0
                
                sSQL = "select sum(lekpr * bestvor) as DATN  from bestrest "
                sSQL = sSQL & " Where linr = " & cLinr
                
                Set rsUW = gdBase.OpenRecordset(sSQL)
                If Not rsUW.EOF Then
                    If Not IsNull(rsUW!DATN) Then
                        dWertLEK = rsUW!DATN
                    End If
                End If
                rsUW.Close
                
                sSQL = "Update BESTLIN set INBESTLEK = '" & dWertLEK & "'"
                sSQL = sSQL & " where linr = " & cLinr
                gdBase.Execute sSQL, dbFailOnError
                
                
                
                'datum lBestellung
                lMaxDate = 0
                
                sSQL = "select max(best_datum) as DATN  from bestrest "
                sSQL = sSQL & " Where linr = " & cLinr
                
                Set rsUW = gdBase.OpenRecordset(sSQL)
                If Not rsUW.EOF Then
                    If Not IsNull(rsUW!DATN) Then
                        lMaxDate = rsUW!DATN
                    End If
                End If
                rsUW.Close
                
                sSQL = "Update BESTLIN set MAXDATE = " & lMaxDate
                sSQL = sSQL & " where linr = " & cLinr
                gdBase.Execute sSQL, dbFailOnError
                
                'datum lBestellung
                
                
                
                lsumKB = 0
                sSQL = "select sum(KB) as MAXI from BESTART "
                
                Set rsUW = gdBase.OpenRecordset(sSQL)
                If Not rsUW.EOF Then
                    If Not IsNull(rsUW!maxi) Then
                        lsumKB = rsUW!maxi
                    End If
                End If
                rsUW.Close
                
                sSQL = "Update BESTLIN set anzkb = " & lsumKB
                sSQL = sSQL & " where linr = " & cLinr
                gdBase.Execute sSQL, dbFailOnError
                
                anzeige "normal", lblanzeige & ".." & j, lblanzeige
                j = j + 1
                
'                lsumFART = 0
'                sSQL = "select sum(UBERBEST1) as MAXI from BESTART2 "
'
'                Set rsUW = gdBase.OpenRecordset(sSQL)
'                If Not rsUW.EOF Then
'                    If Not IsNull(rsUW!Maxi) Then
'                        lsumFART = rsUW!Maxi
'                    End If
'                End If
'                rsUW.Close
'
'                sSQL = "Update BESTLIN set ANZARTF = " & lsumFART
'                sSQL = sSQL & " where linr = " & clinr
'                gdBase.Execute sSQL, dbFailOnError
                
            End If
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    sSQL = "Update BESTLIN set INBESTLEK  = 0 "
    sSQL = sSQL & " where INBESTLEK is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BESTLIN set DIFF  = 0 "
    sSQL = sSQL & " where DIFF is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BESTLIN set ANZART  = 0 "
    sSQL = sSQL & " where ANZART is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BESTLIN set ANZKB  = 0 "
    sSQL = sSQL & " where ANZKB is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BESTLIN set vorlek  = 0 "
    sSQL = sSQL & " where vorlek is null "
    gdBase.Execute sSQL, dbFailOnError
    
    LagerwerteschreibenLINRJetzt lblanzeige
    
    'hier die Pennerstück sind ermitteln
    sSQL = "Select * from BESTLIN "
    Set rsUW = gdBase.OpenRecordset(sSQL)
    If Not rsUW.EOF Then
        rsUW.MoveFirst
        Do While Not rsUW.EOF

        If Not IsNull(rsUW!linr) Then

            rsUW.Edit
            rsUW!ANZARTF = PennerStückErmittlungJetzt(CLng(rsUW!linr))
            rsUW.Update

        End If

        rsUW.MoveNext
        Loop
    End If
    rsUW.Close
    
    sSQL = "Update BESTLIN inner join lisrt on lisrt.linr = BESTLIN.linr set "
    sSQL = sSQL & " BESTLIN.liefbez = lisrt.liefbez "
    sSQL = sSQL & " ,BESTLIN.KUERZEL = lisrt.KUERZEL "
    sSQL = sSQL & " ,BESTLIN.AWERT = lisrt.AWERT "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BESTLIN set ADATE = " & CLng(DateValue(Now))
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BESTLIN set AZEIT = '" & TimeValue(Now) & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BESTLIN set AWERT  = 0 "
    sSQL = sSQL & " where AWERT is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BESTLIN set diff = vorlek - AWERT "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update BESTLIN set KAT  = '1'"
    sSQL = sSQL & " where DIFF > 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BESTLIN set KAT  = '2'"
    sSQL = sSQL & " where DIFF <= 0 and ANZKB > 0"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BESTLIN set KAT  = '3'"
    sSQL = sSQL & " where DIFF <= 0 and ANZKB = 0"
    gdBase.Execute sSQL, dbFailOnError

    
    loeschNEW "SF" & srechnertab, gdBase
    
    
    anzeige "normal", "", lblanzeige
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ErmittleBestVorschlag"
    Fehler.gsFehlertext = "Beim Ermitteln der Bestellvorschläge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    Resume Next
End Sub
Public Function ermumsatzTotal(cKunde As String, bNegpreis As Boolean) As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim dErg    As Double
    
    ermumsatzTotal = "0"
    
    If cKunde <> "" Then
        If IsNumeric(cKunde) = True Then
        
            sSQL = "Select sum(Preis) as maxi from Kassjour where kundnr = " & cKunde
            If bNegpreis Then
            
            Else
                sSQL = sSQL & " and preis > 0 "
            End If
            Set rsrs = gdBase.OpenRecordset(sSQL)
            If Not rsrs.EOF Then
                If Not IsNull(rsrs!maxi) Then
                    dErg = CDbl(rsrs!maxi)
                End If
            End If
            rsrs.Close: Set rsrs = Nothing
            
            
            sSQL = "Select sum(Preis) as maxi from KUNDKASS where kundnr = " & cKunde
            If bNegpreis Then

            Else
                sSQL = sSQL & " and preis > 0 "
            End If
            Set rsrs = gdBase.OpenRecordset(sSQL)
            If Not rsrs.EOF Then
                If Not IsNull(rsrs!maxi) Then
                    dErg = dErg + CDbl(rsrs!maxi)
                End If
            End If
            rsrs.Close: Set rsrs = Nothing
        End If
    End If
    
    ermumsatzTotal = Format$(dErg, "#####0.00")
    
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermumsatzTotal"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermertragTotal(cKunde As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    ermertragTotal = "0"
    
    If cKunde <> "" Then
        If IsNumeric(cKunde) = True Then
    
            sSQL = "Select sum((kassjour.preis)-(kassjour.menge * kassjour.ekpr))as maxi from Kassjour where kundnr = " & cKunde
            Set rsrs = gdBase.OpenRecordset(sSQL)
            If Not rsrs.EOF Then
                If Not IsNull(rsrs!maxi) Then
                    ermertragTotal = rsrs!maxi
                End If
            End If
            rsrs.Close: Set rsrs = Nothing
            
        End If
    End If
    
    ermertragTotal = Format$(ermertragTotal, "#####0.00")
    
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermertragTotal"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermKundenkaufnachUMS(sKN As String) As Double
    On Error GoTo LOKAL_ERROR

    Dim sSQL  As String
    Dim rsrs As Recordset
   
    ermKundenkaufnachUMS = 0
    
    sSQL = "Select sum(preis)as maxi from Kassjour where Kundnr = " & sKN
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        
        If Not IsNull(rsrs!maxi) Then
            ermKundenkaufnachUMS = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermKundenkaufnachUMS"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Function ermKundenkaufnachUMSmitZR(sKN As String, sVon As String, sBis As String) As Double
    On Error GoTo LOKAL_ERROR

    Dim sSQL  As String
    Dim rsrs As Recordset
   
    ermKundenkaufnachUMSmitZR = 0
    
    sSQL = "Select sum(preis)as maxi from Kassjour where Kundnr = " & sKN
    sSQL = sSQL & " and adate Between " & sVon & " and " & sBis
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermKundenkaufnachUMSmitZR = rsrs!maxi
        End If
        
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermKundenkaufnachUMSmitZR"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Function ermKundenkaufnachERT(sKN As String) As Double
    On Error GoTo LOKAL_ERROR

    Dim sSQL  As String
    Dim rsrs As Recordset
   
    ermKundenkaufnachERT = 0
    
    sSQL = "select sum((kassjour.preis)-(kassjour.menge * kassjour.ekpr))as maxi from Kassjour where Kundnr = " & sKN
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        
        If Not IsNull(rsrs!maxi) Then
            ermKundenkaufnachERT = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermKundenkaufnachERT"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Function ermBonusTotal(cKunde As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    ermBonusTotal = "0"
    
    If cKunde <> "" Then
        If IsNumeric(cKunde) = True Then
    
            sSQL = "Select bonus as maxi from kunden where kundnr = " & cKunde
            Set rsrs = gdBase.OpenRecordset(sSQL)
            If Not rsrs.EOF Then
                If Not IsNull(rsrs!maxi) Then
                    ermBonusTotal = rsrs!maxi
                End If
            End If
            rsrs.Close: Set rsrs = Nothing
            
        End If
    End If
    ermBonusTotal = Format$(ermBonusTotal, "#####0.00")
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermBonusTotal"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermKundenkaufnachERTmitZR(sKN As String, sVon As String, sBis As String) As Double
    On Error GoTo LOKAL_ERROR

    Dim sSQL  As String
    Dim rsrs As Recordset
   
    ermKundenkaufnachERTmitZR = 0
    
    sSQL = "select sum((kassjour.preis)-(kassjour.menge * kassjour.ekpr))as maxi from Kassjour where Kundnr = " & sKN
    sSQL = sSQL & " and adate Between " & sVon & " and " & sBis
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermKundenkaufnachERTmitZR = rsrs!maxi
        End If
        
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermKundenkaufnachERTmitZR"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Function ermKundenkaufnachDat(sKN As String, sVon As String, sBis As String) As Long
    On Error GoTo LOKAL_ERROR

    Dim sSQL  As String
    Dim rsrs As Recordset
   
    ermKundenkaufnachDat = 0
    
    sSQL = "Select * from Kassjour where Kundnr = " & sKN
    sSQL = sSQL & " and adate Between " & sVon & " and " & sBis
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        
        ermKundenkaufnachDat = rsrs.RecordCount
        
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermKundenkaufnachDat"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function

Public Function ermKundenkaufnachAGN(sKN As String, sAGN As String) As Long
    On Error GoTo LOKAL_ERROR

    Dim sSQL  As String
    Dim rsrs As Recordset
   
    ermKundenkaufnachAGN = 0
    
    sSQL = "Select * from Kassjour where Kundnr = " & sKN
    sSQL = sSQL & sAGN
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        
        ermKundenkaufnachAGN = rsrs.RecordCount
        
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermKundenkaufnachAGN"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Function ermKundenkaufnachLL(sKN As String, sLL As String) As Long
    On Error GoTo LOKAL_ERROR

    Dim sSQL  As String
    Dim rsrs As Recordset
   
    ermKundenkaufnachLL = 0
    
    sSQL = "Select * from Kassjour where Kundnr = " & sKN
    sSQL = sSQL & sLL
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        
        ermKundenkaufnachLL = rsrs.RecordCount
        
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermKundenkaufnachLL"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Function ermKundenkaufnachAGNmitZR(sKN As String, sAGN As String, sVon As String, sBis As String) As Long
    On Error GoTo LOKAL_ERROR

    Dim sSQL  As String
    Dim rsrs As Recordset
   
    ermKundenkaufnachAGNmitZR = 0
    
    sSQL = "Select * from Kassjour where Kundnr = " & sKN
    sSQL = sSQL & sAGN
    sSQL = sSQL & " and adate Between " & sVon & " and " & sBis
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        
        ermKundenkaufnachAGNmitZR = rsrs.RecordCount
        
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermKundenkaufnachAGNmitZR"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Function ermKundenkaufnachLLmitZR(sKN As String, sLL As String, sVon As String, sBis As String) As Long
    On Error GoTo LOKAL_ERROR

    Dim sSQL  As String
    Dim rsrs As Recordset
   
    ermKundenkaufnachLLmitZR = 0
    
    sSQL = "Select * from Kassjour where Kundnr = " & sKN
    sSQL = sSQL & sLL
    sSQL = sSQL & " and adate Between " & sVon & " and " & sBis
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        
        ermKundenkaufnachLLmitZR = rsrs.RecordCount
        
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermKundenkaufnachLLmitZR"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Function ermKundenkauf(sKN As String) As Long
    On Error GoTo LOKAL_ERROR

    Dim sSQL  As String
    Dim rsrs As Recordset
   
     
    ermKundenkauf = 0
    
    sSQL = "Select * from Kassjour where Kundnr = " & sKN
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        
        ermKundenkauf = rsrs.RecordCount
        
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermKundenkauf"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Function inWorten$(ByRef Wert As String)
Const Blöcke = 4
'max Anzahl von Dreierblöcken in einer Zahl (z.B. 4 = max bis 999 999 999 999)
Dim Block$(Blöcke)
Dim Text$(Blöcke)
Dim Gruppe$(Blöcke)
Dim GrEndSg$(Blöcke)
Dim GrEndPl$(Blöcke)
Dim Einer$(10)
Dim Einer2$(10)
Einer$(0) = ""
Einer$(1) = "eins"
Einer$(2) = "zwei"
Einer$(3) = "drei"
Einer$(4) = "vier"
Einer$(5) = "fünf"
Einer$(6) = "sechs"
Einer$(7) = "sieben"
Einer$(8) = "acht"
Einer$(9) = "neun"
Einer2$(0) = ""
Einer2$(1) = "ein"
Einer2$(2) = "zwei"
Einer2$(3) = "drei"
Einer2$(4) = "vier"
Einer2$(5) = "fünf"
Einer2$(6) = "sech"
Einer2$(7) = "sieb"
Einer2$(8) = "acht"
Einer2$(9) = "neun"
Gruppe$(1) = ""
Gruppe$(2) = "tausend"
Gruppe$(3) = " Million"
Gruppe$(4) = " Milliarde"
' Gruppenendung Singular
GrEndSg$(1) = ""
GrEndSg$(2) = ""
GrEndSg$(3) = " "
GrEndSg$(4) = " "
' Gruppenendung Plural
GrEndPl$(1) = ""
GrEndPl$(2) = ""
GrEndPl$(3) = "en "
GrEndPl$(4) = "n "
Dim i As Integer
For i = 1 To Blöcke
Block$(i) = ""
Text$(i) = ""
Next
'**************************************************************************
'* Alle Punkte entfernen
'**************************************************************************
Dim Pos As Long
Dim NK As String
Dim TextG As String

Pos = InStr(Wert$, ".")
While Pos > 0
Wert$ = Left$(Wert$, Pos - 1) + Right$(Wert$, Len(Wert$) - Pos)
Pos = InStr(Pos, Wert$, ".")
Wend
'**************************************************************************
'* Nachkommastellen NK$ schreiben
'**************************************************************************
Pos = InStr(Wert$, ",")
If Pos > 0 Then

NK$ = Right$(Wert$, Len(Wert$) - Pos)
Wert$ = Left$(Wert$, Pos - 1)
Else
NK$ = ""
End If

For i = 1 To Blöcke
If Len(Wert$) > 3 Then
Block$(i) = Right$(Wert$, 3)
Wert$ = Left$(Wert$, Len(Wert$) - 3)
Else
Block$(i) = Wert$
Wert$ = ""
End If
If Block$(i) <> "" Then
If Len(Block$(i)) = 3 Then
If Block$(i) = "000" Then
Text$(i) = ""
ElseIf Left$(Block$(i), 1) = "1" Then
Text$(i) = "einhundert"
ElseIf Left$(Block$(i), 1) = "0" Then
Text$(i) = ""
Else
Text$(i) = Text$(i) + Einer$(Val(Left$(Block$(i), 1))) + "hundert"
End If
Block$(i) = Right$(Block$(i), 2)
End If

If Len(Block$(i)) = 2 Then
If Left$(Block$(i), 1) = "0" Then
Text$(i) = Text$(i) + Einer$(Val(Right$(Block$(i), 1)))
ElseIf Left$(Block$(i), 1) = "1" Then
If Left$(Block$(i), 2) = "11" Then
Text$(i) = Text$(i) + "elf"
ElseIf Left$(Block$(i), 2) = "12" Then
Text$(i) = Text$(i) + "zwölf"
Else
Text$(i) = Text$(i) + Einer2$(Val(Right$(Block$(i), 1))) + "zehn"
End If
ElseIf Left$(Block$(i), 1) = "2" Then
If Left$(Block$(i), 2) = "21" Then
Text$(i) = Text$(i) + "ein"
Else
Text$(i) = Text$(i) + Einer$(Val(Right$(Block$(i), 1)))
End If
If Left$(Block$(i), 2) <> "20" Then
Text$(i) = Text$(i) + "und"
End If
Text$(i) = Text$(i) + "zwanzig"
ElseIf Left$(Block$(i), 1) = "3" Then
If Left$(Block$(i), 2) = "31" Then
Text$(i) = Text$(i) + "ein"
Else
Text$(i) = Text$(i) + Einer$(Val(Right$(Block$(i), 1)))
End If
If Left$(Block$(i), 2) <> "30" Then
Text$(i) = Text$(i) + "und"
End If
Text$(i) = Text$(i) + "dreißig"
Else
If Right$(Block$(i), 1) = "1" Then
Text$(i) = Text$(i) + "ein"
Else
Text$(i) = Text$(i) + Einer$(Val(Right$(Block$(i), 1)))
End If
If Right$(Block$(i), 1) <> "0" Then
Text$(i) = Text$(i) + "und"
End If
Text$(i) = Text$(i) + Einer2$(Val(Left$(Block$(i), 1))) + "zig"
End If
End If
If Len(Block$(i)) = 1 Then
Text$(i) = Text$(i) + Einer$(Val(Right$(Block$(i), 1)))
End If
End If
If Text$(i) <> "" Then
End If
Next
For i = Blöcke To 1 Step -1
If Text$(i) <> "" Then
If Text$(i) = "eins" Then
If i > 2 Then
Text$(i) = "eine"
ElseIf i = 2 Then
Text$(i) = "ein"
End If
Text$(i) = Text$(i) + Gruppe$(i)
Text$(i) = Text$(i) + GrEndSg$(i)
Else
Text$(i) = Text$(i) + Gruppe$(i)
Text$(i) = Text$(i) + GrEndPl$(i)
End If
End If
TextG$ = TextG$ + Text$(i)
Next
If TextG$ = "" Then
TextG$ = "null"
End If
If (NK$ <> "") And (NK$ <> "0") And (NK$ <> "00") Then
If Len(NK$) = 1 Then
NK$ = NK$ + "0"
End If
TextG$ = TextG$ + " und " + NK$ + "/100"
End If
' TextG$ = Chr$(Asc(Left$(TextG$, 1)) - 32) + Right$(TextG$, Len(TextG$) - 1)
inWorten$ = TextG$
End Function

Public Function LeseVerkäufeKundeTotal(cSuch As String) As String
    On Error GoTo LOKAL_ERROR
    
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    LeseVerkäufeKundeTotal = "0"
    
    If cSuch <> "" Then
        If IsNumeric(cSuch) = True Then
    
            cSQL = "Select sum(Preis) as UMS from Kassjour where KUNDNR = " & cSuch
            Set rsrs = gdBase.OpenRecordset(cSQL)
            If Not rsrs.EOF Then
                If Not IsNull(rsrs!UMS) Then
                    LeseVerkäufeKundeTotal = rsrs!UMS
                End If
            End If
            rsrs.Close: Set rsrs = Nothing
        End If
    End If
    
    LeseVerkäufeKundeTotal = Format$(LeseVerkäufeKundeTotal, "#####0.00")
    

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "LeseVerkäufeKundeTotal"
    Fehler.gsFehlertext = "Im Programmteil Kunden bearbeiten ist ein Fehler aufgetreten."
    
    
    Fehlermeldung1
End Function
Public Function Kundevorhanden(cKunde As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rs As Recordset
    
    Kundevorhanden = False
    
    sSQL = "select * from kunden where kundnr = " & cKunde
    sSQL = sSQL & " and (STATUS <> 'D' or STATUS is null)"
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
        Kundevorhanden = True
    End If
    rs.Close: Set rs = Nothing
    
    Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "Kundevorhanden"
    Fehler.gsFehlertext = "Beim Ermitteln der Farbe ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub Insertgrolief(cLinr As String)
On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
   
    If gbSQLSERVER = True Then
        cSQL = "Delete from grolief where linr = " & cLinr
    Else
        cSQL = "Delete from grolief where linr = " & cLinr
    End If
    
    gdBase.Execute cSQL, dbFailOnError + dbSQLPassThrough
    
    cSQL = "Insert into grolief  (linr) values (" & cLinr & ")"
    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "Insertgrolief"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub diefarbezeichnung(stabf As String, dabx As Database)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String


    sSQL = "Update " & stabf & " set farbtext = 'neue Artikel' "
    sSQL = sSQL & "  where farbnr = 98 "
    dabx.Execute sSQL, dbFailOnError
    
    sSQL = "Update " & stabf & " set farbtext = 'nicht geliefert' "
    sSQL = sSQL & "  where farbnr = 95 "
    dabx.Execute sSQL, dbFailOnError
    
    sSQL = "Update " & stabf & "  set farbtext = 'für Preisaktion vorgesehen' "
    sSQL = sSQL & "  where farbnr = 94 "
    dabx.Execute sSQL, dbFailOnError
    
    sSQL = "Update " & stabf & "  set farbtext = 'befinden sich in Preisaktion' "
    sSQL = sSQL & "  where farbnr = 93 "
    dabx.Execute sSQL, dbFailOnError
    
    sSQL = "Update " & stabf & "  set farbtext = 'seit 2 Jahren oder nie verkauft' "
    sSQL = sSQL & "  where farbnr = 92 "
    dabx.Execute sSQL, dbFailOnError
    
    
    
    sSQL = "Update " & stabf & "  inner join Farbmerk on " & stabf & ".farbnr = farbmerk.farbnr"
    sSQL = sSQL & " set " & stabf & ".farbtext = farbmerk.farbtext "
    dabx.Execute sSQL, dbFailOnError
    
   
    sSQL = "Update " & stabf & "  set farbtext = 'ohne Beschreibung' "
    sSQL = sSQL & "  where farbtext is null "
    dabx.Execute sSQL, dbFailOnError
    
    
       
    Dim i As Integer
    Dim lFarbert As Long
    Dim lFarbert2 As Long
    
    For i = 1 To 9
    lFarbert = CDec(glfarbe(i))
    lFarbert2 = vbBlack

    
        sSQL = "Update " & stabf & "  set farbwert =  " & lFarbert
        sSQL = sSQL & " , farbwerts = " & lFarbert2
        sSQL = sSQL & "  where farbnr =  " & i
        
        dabx.Execute sSQL, dbFailOnError
        
    
    Next i
    
    For i = 1 To 9
    lFarbert = CDec(glfarbe2(i))
    lFarbert2 = vbBlack

    
        sSQL = "Update " & stabf & "  set farbwert =  " & lFarbert
        sSQL = sSQL & " , farbwerts = " & lFarbert2
        sSQL = sSQL & "  where farbnr =  " & i + 10
        
        dabx.Execute sSQL, dbFailOnError
        
    
    Next i
    
    lFarbert = vbRed
    lFarbert2 = vbWhite

    
    sSQL = "Update " & stabf & " set farbwerts =  " & lFarbert
    sSQL = sSQL & " , farbwert = " & lFarbert2
    sSQL = sSQL & "  where farbnr =  98 "
    dabx.Execute sSQL, dbFailOnError
    
    lFarbert = vbBlack
    lFarbert2 = vbBlue

    
    sSQL = "Update " & stabf & "  set farbwerts =  " & lFarbert
    sSQL = sSQL & " , farbwert = " & lFarbert2
    sSQL = sSQL & "  where farbnr =  95 "
    dabx.Execute sSQL, dbFailOnError
    
    lFarbert = vbWhite
    lFarbert2 = vbBlack
    
    sSQL = "Update " & stabf & "  set farbwerts =  " & lFarbert
    sSQL = sSQL & " , farbwert = " & lFarbert2
    sSQL = sSQL & "  where farbnr =  92 "
    dabx.Execute sSQL, dbFailOnError
    
    lFarbert = vbGreen
    lFarbert2 = vbWhite
    
    sSQL = "Update " & stabf & "  set farbwerts =  " & lFarbert
    sSQL = sSQL & " , farbwert = " & lFarbert2
    sSQL = sSQL & "  where farbnr =  93 "
    dabx.Execute sSQL, dbFailOnError
    
    lFarbert = vbBlue
    lFarbert2 = glfarbe(0)
    
    sSQL = "Update " & stabf & "  set farbwerts =  " & lFarbert
    sSQL = sSQL & " , farbwert = " & lFarbert2
    sSQL = sSQL & "  where farbnr =  94 "
    dabx.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "diefarbezeichnung"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1

End Sub

Public Sub delgrolief(cLinr As String)
On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
   
    cSQL = "Delete from grolief where linr = " & cLinr
    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "delgrolief"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function KundenbestBestätigung(cSuch As String, dMenge As Double) As Boolean
    On Error GoTo LOKAL_ERROR
    
    '1. offene Kundenbestellungen ermitteln
    '2. Kundenbestellungen als geliefert bestätigen
    '3. eventuellen Rest stehen lassen
    
    Dim cSQL            As String
    Dim rsrs            As Recordset
    Dim lMengeGelief    As Long
    Dim LmengeUebrig    As Long
    Dim lMengeBest      As Long
    Dim lMengeZuteil    As Long
    
    lMengeGelief = CLng(dMenge)
    
    KundenbestBestätigung = False
    
    cSQL = "Select * from KUNDBEST where artnr = " & Val(cSuch)
    cSQL = cSQL & " and  (StatusARTIKEL = 'INBESTELLUNG' "
    cSQL = cSQL & " or StatusARTIKEL = 'BESTELLT')"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        KundenbestBestätigung = True
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If KundenbestBestätigung = False Then
        Exit Function
    End If
    
    cSQL = "Select * from KUNDBEST where artnr = " & Val(cSuch)
    cSQL = cSQL & " and  (StatusARTIKEL = 'INBESTELLUNG' "
    cSQL = cSQL & " or StatusARTIKEL = 'BESTELLT') order by bestelltam asc"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            
            If Not IsNull(rsrs!BestelltMenge) Then
                lMengeBest = CLng(rsrs!BestelltMenge)
            End If
            
            If lMengeGelief >= lMengeBest Then
            
                lMengeGelief = lMengeGelief - lMengeBest
                lMengeZuteil = lMengeBest
                
                rsrs.Edit
                rsrs!StatusARTIKEL = "GELIEFERT"
                rsrs!statusKunde = "INFORMIEREN"
                rsrs.Update
                
            Else
                If lMengeGelief > 0 Then
                    lMengeZuteil = lMengeGelief
                    LmengeUebrig = lMengeBest - lMengeGelief
                    lMengeGelief = lMengeGelief - lMengeBest
                    
                    InsertKuBest cSuch, rsrs!Bestelltam, rsrs!Bestelltum, LmengeUebrig
                    
                    rsrs.Edit
                    rsrs!BestelltPreis = (CLng(rsrs!BestelltPreis) / CLng(rsrs!BestelltMenge)) * lMengeZuteil
                    rsrs!BestelltMenge = lMengeZuteil
                    rsrs!StatusARTIKEL = "GELIEFERT"
                    rsrs!statusKunde = "INFORMIEREN"
                    rsrs.Update
                    
                Else
                
                    rsrs.Close: Set rsrs = Nothing
                    Exit Function
                
                End If
            End If
    
        rsrs.MoveNext
        Loop
        
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "KundenbestBestätigung"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Function
Public Sub InsertKuBest(sArtnr As String, cBestam As String, cBestum As String, lMenge As Long)
On Error GoTo LOKAL_ERROR

Dim sSQL As String

loeschNEW "KUTERT", gdBase

sSQL = "Select "
sSQL = sSQL & " ARTNR "
sSQL = sSQL & ", BEZEICH  "
sSQL = sSQL & ", KUNDNR  "
sSQL = sSQL & ", BEDNU  "
sSQL = sSQL & ", EKPR "
sSQL = sSQL & ", VKPR "
sSQL = sSQL & ", MWST "
sSQL = sSQL & ", FARBE "
sSQL = sSQL & ", FARBTEXT "
sSQL = sSQL & ", Filiale "
sSQL = sSQL & ", SENDOK  "
sSQL = sSQL & ", STATUSARTIKEL "
sSQL = sSQL & ", 'INBEA' as STATUSKUNDE "
sSQL = sSQL & ", BESTELLTAM  "
sSQL = sSQL & ", BESTELLTUM  "
sSQL = sSQL & ", BESTELLTPREIS  "
sSQL = sSQL & ", BESTELLTMENGE  "
sSQL = sSQL & " into KUTERT  from Kundbest "
sSQL = sSQL & " where artnr = " & sArtnr
sSQL = sSQL & " and BESTELLTAM = " & CLng(DateValue(cBestam))
sSQL = sSQL & " and BESTELLTUM  = '" & cBestum & "'"
sSQL = sSQL & " and  (StatusARTIKEL = 'INBESTELLUNG' "
sSQL = sSQL & " or StatusARTIKEL = 'BESTELLT') "
gdBase.Execute sSQL, dbFailOnError

sSQL = "Update KUTERT Set "
sSQL = sSQL & " BESTELLTPREIS = (BESTELLTPREIS/BESTELLTmenge)* " & lMenge
sSQL = sSQL & ", BESTELLTMENGE = " & lMenge
sSQL = sSQL & ", StatusKUNDE = '' "
sSQL = sSQL & " where StatusKUNDE = 'INBEA' "
gdBase.Execute sSQL, dbFailOnError



sSQL = "Insert into Kundbest  Select * "
sSQL = sSQL & " from KUTERT "
gdBase.Execute sSQL, dbFailOnError

loeschNEW "KUTERT", gdBase

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "InsertKuBest"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub UpdateKuBestKUNDENSTATUS(sStatusKundealt As String, sStatusKundeneu As String)
On Error GoTo LOKAL_ERROR

Dim sSQL As String

sSQL = "Update Kundbest set StatusKunde =  '" & sStatusKundeneu & "'"
sSQL = sSQL & " Where StatusKunde =  '" & sStatusKundealt & "'"
gdBase.Execute sSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "UpdateKuBestKUNDENSTATUS"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub KB(sArtikelstatus As String, sKundenstatus As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String

    loeschNEW "KUOKB", gdBase
    CreateTable "KUOKB", gdBase
    
    sSQL = "Insert into KUOKB Select "
    sSQL = sSQL & " ARTNR "
    sSQL = sSQL & ", BEZEICH"
    sSQL = sSQL & ", BEDNU  "
    sSQL = sSQL & ", EKPR "
    sSQL = sSQL & ", VKPR "
    sSQL = sSQL & ", MWST"
    sSQL = sSQL & ", FARBE "
    sSQL = sSQL & ", FARBTEXT "
    sSQL = sSQL & ", Filiale "
    sSQL = sSQL & ", SENDOK "
    sSQL = sSQL & ", STATUSARTIKEL "
    sSQL = sSQL & ", STATUSKUNDE "
    sSQL = sSQL & ", BESTELLTAM  "
    sSQL = sSQL & ", BESTELLTUM  "
    sSQL = sSQL & ", BESTELLTPREIS  "
    sSQL = sSQL & ", BESTELLTMENGE  "
    sSQL = sSQL & ", KUNDNR "
    sSQL = sSQL & "  from KUNDBEST where STATUSARTIKEL = '" & sArtikelstatus & "' "
    
    If sKundenstatus <> "" Then
        sSQL = sSQL & "  and  STATUSKunde = '" & sKundenstatus & "' "
    End If
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KUOKB inner join KUNDEN on KUOKB.KUNDNR = KUNDEN.KUNDNR"
    sSQL = sSQL & " SET KUOKB.TEL = KUNDEN.TEL "
    sSQL = sSQL & ", KUOKB.FAXNR = KUNDEN.FAXNR "
    sSQL = sSQL & ", KUOKB.EMAIL = KUNDEN.EMAIL "
    sSQL = sSQL & ", KUOKB.MOBILTEL = KUNDEN.MOBILTEL "
    sSQL = sSQL & ", KUOKB.VORNAME = KUNDEN.VORNAME "
    
    sSQL = sSQL & ", KUOKB.NAME = KUNDEN.NAME "
    sSQL = sSQL & ", KUOKB.STRASSE = KUNDEN.STRASSE "
    sSQL = sSQL & ", KUOKB.PLZ = KUNDEN.PLZ "
    sSQL = sSQL & ", KUOKB.ORT = KUNDEN.STADT "
    sSQL = sSQL & ", KUOKB.TITEL = KUNDEN.TITEL "
    sSQL = sSQL & ", KUOKB.FIRMA = KUNDEN.FIRMA "
    gdBase.Execute sSQL, dbFailOnError
      
    Select Case sArtikelstatus
        Case "INBESTELLUNG"
            reportbildschirm "", "aWKL77a"
        Case "BESTELLT"
            reportbildschirm "", "aWKL77b"
        Case "GELIEFERT"
            reportbildschirm "", "aWKL77c"
        Case "NICHTGELIEFERT"
            reportbildschirm "", "aWKL77d"
    End Select
    
    Pause (2)
    loeschNEW "KUOKB", gdBase

Exit Sub
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "KB"
    Fehler.gsFehlertext = "Im Programmteil Kundenlistengenerator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
  
End Sub
Public Function WasIstInhaltBez(sArtnr As String) As String
On Error GoTo LOKAL_ERROR

Dim sSQL As String
Dim rsrs As Recordset

WasIstInhaltBez = ""

sArtnr = SwapStr(sArtnr, "+", "")
sArtnr = SwapStr(sArtnr, "-", "")

If IsNumeric(sArtnr) Then

    sSQL = "Select INHALTBEZ from ARTIKEL where ARTNR = " & sArtnr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!INHALTBEZ) Then
            WasIstInhaltBez = UCase(rsrs!INHALTBEZ)
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
End If

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "WasIstInhaltBez"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function Speicherblock(sArt As String, sHauptext As String, sKurztext As String) As Boolean
On Error GoTo LOKAL_ERROR

Dim sSQL As String
Dim rsrs As Recordset
Dim iRet As Integer

Speicherblock = False

If sHauptext = "" Then
    Speicherblock = True
    Exit Function
End If

If sKurztext = "" Then
    sKurztext = Left(sHauptext, 35)
End If

sSQL = "Select * from TEXTBLOCK where Kurzbeschreib = '" & sKurztext & "'"
sSQL = sSQL & " and TEXTART = '" & sArt & "'"
Set rsrs = gdBase.OpenRecordset(sSQL)

If Not rsrs.EOF Then
    iRet = MsgBox("Ein Text mit dieser Kurzbezeichnung ist schon vorhanden, überscheiben?", vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
    If iRet = vbNo Then
        Speicherblock = True
        Exit Function
    End If
    rsrs.Edit
Else
    rsrs.AddNew
End If

rsrs!kurzbeschreib = sKurztext
rsrs!BESCHREIB = sHauptext
rsrs!TEXTART = sArt

BeginTrans
    rsrs.Update
CommitTrans

rsrs.Close: Set rsrs = Nothing

Speicherblock = True

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "Speicherblock"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Function
Public Sub ZeigeblockinList(sArt As String, Listx As ListBox)
On Error GoTo LOKAL_ERROR

Dim sSQL    As String
Dim rsrs    As Recordset
Dim cSatz   As String
Listx.Clear

sSQL = "Select * from TEXTBLOCK  "
sSQL = sSQL & " where TEXTART = '" & sArt & "'"
Set rsrs = gdBase.OpenRecordset(sSQL)

If Not rsrs.EOF Then
    rsrs.MoveFirst
    Do While Not rsrs.EOF
    
    cSatz = ""
    If Not IsNull(rsrs!kurzbeschreib) Then
        cSatz = rsrs!kurzbeschreib & Space(100 - Len(rsrs!kurzbeschreib))
        cSatz = cSatz & rsrs!Tbnr
        Listx.AddItem cSatz
    End If
    
    rsrs.MoveNext
    Loop
    
End If

rsrs.Close: Set rsrs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ZeigeblockinList"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub ZeigeblockinEinzelteile(sArt As String, Listx As ListBox, t1 As TextBox, t2 As TextBox, lTBNR As Long)
On Error GoTo LOKAL_ERROR

Dim sSQL    As String
Dim rsrs    As Recordset


sSQL = "Select * from TEXTBLOCK  "
sSQL = sSQL & " where TBNR = " & lTBNR
Set rsrs = gdBase.OpenRecordset(sSQL)

If Not rsrs.EOF Then
    If Not IsNull(rsrs!kurzbeschreib) Then
        t1.Text = rsrs!kurzbeschreib
    End If

    If Not IsNull(rsrs!BESCHREIB) Then
        t2.Text = rsrs!BESCHREIB
    End If
End If

rsrs.Close: Set rsrs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ZeigeblockinEinzelteile"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub DELblock(lTBNR As Long)
On Error GoTo LOKAL_ERROR

Dim sSQL As String

sSQL = "Delete from TEXTBLOCK where TBNR = " & lTBNR
gdBase.Execute sSQL, dbFailOnError


Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "DELblock"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function ErmlzZugangM(cART As String) As String
    On Error GoTo LOKAL_ERROR
    
    ErmlzZugangM = "0"
    
    Dim cSQL As String
    Dim rsINB As Recordset
    
    Dim cSQL1 As String
    Dim rsINB1 As Recordset
    
    cSQL = "Select max(adate) as maxdate  from Zugang where ARTNR = " & cART & " "
    Set rsINB = gdBase.OpenRecordset(cSQL)
    If Not rsINB.EOF Then
        If Not IsNull(rsINB!MaxDate) Then
            ErmlzZugangM = rsINB!MaxDate
            
            cSQL1 = "Select sum(bewegung) as bew  from Zugang where ARTNR = " & cART & " "
            cSQL1 = cSQL1 & " and adate = " & CLng(DateValue(ErmlzZugangM))
            Set rsINB1 = gdBase.OpenRecordset(cSQL1)
            If Not rsINB1.EOF Then
                If Not IsNull(rsINB1!bEW) Then
                    ErmlzZugangM = ErmlzZugangM & "(" & rsINB1!bEW & ")"
                
                End If
            End If
            rsINB1.Close: Set rsINB1 = Nothing
        
        End If
    End If
    rsINB.Close: Set rsINB = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ErmlzZugangM"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Sub ErmlzDreiZugaenge(cART As String, Listx As ListBox)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsINB   As Recordset
    
    Dim cSQL1   As String
    Dim cSQL2   As String
    Dim rsINB1  As Recordset
    Dim rsINB2  As Recordset
    Dim sErg    As String
    
    Dim sZugang As String
    
    Dim sKurz   As String
    
    Listx.Clear
    
    cSQL = "Select distinct(adate) as maxdate  from Zugang where ARTNR = " & cART & " order by adate desc "
    Set rsINB = gdBase.OpenRecordset(cSQL)
    If Not rsINB.EOF Then
        Do While Not rsINB.EOF
            If Not IsNull(rsINB!MaxDate) Then
                sErg = rsINB!MaxDate
                
                cSQL1 = "Select sum(bewegung) as bew  from Zugang where ARTNR = " & cART & " "
                cSQL1 = cSQL1 & " and adate = " & CLng(DateValue(sErg))
                Set rsINB1 = gdBase.OpenRecordset(cSQL1)
                If Not rsINB1.EOF Then
                    If Not IsNull(rsINB1!bEW) Then
                    
                        sZugang = rsINB1!bEW
                    
                        cSQL2 = "Select max(linr) as maxlinr  from Zugang where ARTNR = " & cART & " "
                        cSQL2 = cSQL2 & " and adate = " & CLng(DateValue(sErg))
                        Set rsINB2 = gdBase.OpenRecordset(cSQL2)
                        If Not rsINB2.EOF Then
                            If Not IsNull(rsINB2!maxlinr) Then
                            
                                sKurz = rsINB2!maxlinr
                                sKurz = ermLiefKürzelmitLiefvorgabe(cART, CLng(sKurz))
                                
                                sErg = sErg & "(" & sZugang & ") " & sKurz
                                Listx.AddItem sErg
                            End If
                        End If
                        rsINB2.Close: Set rsINB2 = Nothing
                    
                        
                    End If
                End If
                rsINB1.Close: Set rsINB1 = Nothing
            
            End If
            rsINB.MoveNext
        Loop
    End If
    rsINB.Close: Set rsINB = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ErmlzDreiZugaenge"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Public Function ErmlzVK(cART As String) As String
    On Error GoTo LOKAL_ERROR
    
    ErmlzVK = "0"
    
    Dim cSQL As String
    Dim rsINB As Recordset
    
    cSQL = "Select max(adate) as maxdate from Kassjour where ARTNR = " & cART & "  "
    Set rsINB = gdBase.OpenRecordset(cSQL)
    If Not rsINB.EOF Then
        If Not IsNull(rsINB!MaxDate) Then
            ErmlzVK = rsINB!MaxDate
        End If
    
    End If
    rsINB.Close: Set rsINB = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ErmlzVK"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function

Public Function SchnittEKBerechnung(sArt As String, lLinr As Long, zugang As Long, RechnungsEK As Double, _
bestandalt As Long) As Double
On Error GoTo LOKAL_ERROR

    Dim sSQL            As String
    Dim WertAlt         As Double
    Dim SchnittEk       As Double
    Dim ListenEk        As Double
    Dim WertZugang      As Double
    Dim rsArt           As Recordset
    Dim Anzahlgesamt    As Long
    Dim Wertgesamt      As Double
    Dim cekalt1         As String
    
    
    SchnittEKBerechnung = 0
    
    SchnittEk = 0
    
    
    sSQL = "Select * from Artikel where artnr = " & sArt
    Set rsArt = gdBase.OpenRecordset(sSQL)
    If Not rsArt.EOF Then
         If Not IsNull(rsArt!ekpr) Then
            cekalt1 = rsArt!ekpr
            If IsNumeric(cekalt1) Then
                SchnittEk = CDbl(cekalt1)
            Else
                SchnittEk = 0
            End If
            
        Else
            SchnittEk = 0
        End If
    End If
    rsArt.Close: Set rsArt = Nothing
    
    
    'Ist bisheriger Schnittek = 0 oder kleiner?
    If SchnittEk <= 0 Then
    
        ListenEk = 0
        sSQL = "Select * from artlief where artnr = " & sArt
        sSQL = sSQL & " and LINR = " & lLinr
        Set rsArt = gdBase.OpenRecordset(sSQL)
        If Not rsArt.EOF Then
            If Not IsNull(rsArt!lekpr) Then
                ListenEk = rsArt!lekpr
            End If
        End If
        rsArt.Close: Set rsArt = Nothing
        
        SchnittEk = ListenEk
        
        If SchnittEk < 0 Then SchnittEk = 0
    
    End If
    
    
    If zugang <> 0 Then
    
        'ist bisheriger Bestand < 0 ?
        If bestandalt < 1 Then bestandalt = 0
        
        
        WertAlt = bestandalt * SchnittEk
        
'        If zugang < 1 Then zugang = 0
'        zugang = -1

        
        WertZugang = zugang * RechnungsEK
        
        Anzahlgesamt = zugang + bestandalt
        
        
        Wertgesamt = CDbl(Format(WertAlt, "#####0.00")) + CDbl(Format(WertZugang, "#####0.00"))
        
        
        If Anzahlgesamt = 0 Then
            SchnittEKBerechnung = SchnittEk
        Else
            SchnittEKBerechnung = Wertgesamt / Anzahlgesamt
        End If
    Else
        SchnittEKBerechnung = SchnittEk
    End If
    
    Exit Function
LOKAL_ERROR:
'    If err.Number = 6 Then
'
'        SchnittEKBerechnung = ListenEk
'        If SchnittEKBerechnung < 0 Then SchnittEKBerechnung = 0
'        Exit Function
'
'    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul1"
        Fehler.gsFunktion = "SchnittEKBerechnung"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten. Artikel: " & sArt & " Bestand alt: " & bestandalt & " Zugang: " & zugang
        
        Fehlermeldung1
'    End If
    
End Function
Public Function HoleNaechsteReNr() As String
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim lrenr As Long
    
    HoleNaechsteReNr = ""
    
    cSQL = "Select max(val(RENR)) as MAXRENR from REKOPF "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!MAXRENR) Then
            HoleNaechsteReNr = rsrs!MAXRENR
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    lrenr = Val(HoleNaechsteReNr)
    lrenr = lrenr + 1
    HoleNaechsteReNr = Trim$(Str$(lrenr))
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "HoleNaechsteReNr"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function

Public Sub SchubladeOeffnen()
    On Error GoTo LOKAL_ERROR
    
    Dim dAktZeit        As Double
    Dim dNeuZeit        As Double
    Dim lRet            As Long
    Dim lDatum          As Long
    Dim cSQL            As String
    Dim aDeviceName     As String
    Dim cEscapeSequenz  As String
    Dim rsrs            As Recordset
    
    setzedrucker gcBonDrucker
    
    lDatum = Fix(Now)                   'Drucker an, Display aus, Init Drucker
    aDeviceName = Printer.DeviceName
    cEscapeSequenz = gcInit
    OpenDrawer aDeviceName, cEscapeSequenz
    
    If gbLadeCom Then
        OpenDrawerViaComPortModul20
    Else
    
        If gbAPI = False Then
            dAktZeit = Time
            lRet = Shell("Command.com /C " & gcPfad & "LADE.EXE", 6)
            dNeuZeit = Time
            Do While dNeuZeit - dAktZeit < (2 / 86400)
                dNeuZeit = Time
            Loop
        Else
            aDeviceName = Printer.DeviceName
            cEscapeSequenz = gcLade
            OpenDrawer aDeviceName, cEscapeSequenz
        End If
        
    End If
    cSQL = "Select ADATE, KASNUM, GELDFACH from AFCSTAT where ADATE = " & Trim$(Str$(lDatum)) & " and KASNUM = " & gcKasNum & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If rsrs.EOF Then
        rsrs.AddNew
    Else
        rsrs.Edit
    End If
    
    'Datum und Kassennummer setzen
    rsrs!ADATE = lDatum
    rsrs!kasnum = Val(gcKasNum)
    
    If IsNull(rsrs!GELDFACH) = True Then
        rsrs!GELDFACH = 1
    Else
        rsrs!GELDFACH = rsrs!GELDFACH + 1
    End If
    
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "SchubladeOeffnen"
    Fehler.gsFehlertext = "Beim Schublade öffnen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function NEWfnCheck4UpdateDateiWKL00(bFromStart As Boolean) As Integer
    On Error GoTo LOKAL_ERROR
    
    Dim iVersion        As Integer
    Dim iKillVersion    As Integer
    Dim i               As Integer
    Dim lFileSize       As Long
    Dim iVersionEnd     As Integer
    Dim lZähler         As Long
    
    NEWfnCheck4UpdateDateiWKL00 = 0
    
    If bFromStart = True Then
        frmWKL00.txtStatus.Text = 0
    End If
    
    iVersionEnd = WKVersion
    iVersionEnd = iVersionEnd + 1000
    
    iVersion = WKVersion
    iVersion = iVersion + 1
    
    iKillVersion = WKVersion - 1
    
    lZähler = 0
    For i = iKillVersion To 1500 Step -1
    

        lZähler = lZähler + 1
        
        If bFromStart = True Then
        
            If lZähler >= 1000 Then
                lZähler = 0
            End If
            frmWKL00.txtStatus.Text = lZähler / 10
        End If
        
        Kill gsUpdPfad & "\WK" & i & ".lzh"
        Kill gsUpdPfad & "\REPO" & i & ".lzh"
    Next i
    
    If bFromStart = True Then
        frmWKL00.txtStatus.Text = 33
    End If
    
    lZähler = 0
    
    For i = iVersion To iVersionEnd
        lZähler = lZähler + 1
        
        If bFromStart = True Then
        
            If lZähler >= 1000 Then
                lZähler = 0
            End If
            frmWKL00.txtStatus.Text = lZähler / 10
        End If
    
        If Modul6.FindFile(gsUpdPfad, "WK" & i & ".lzh") Then
            gsUpdDatName = "WK" & i & ".lzh"
            lFileSize = fnFileSize(gsUpdPfad & "\" & gsUpdDatName)
            If lFileSize > 2400000 Then
                NEWfnCheck4UpdateDateiWKL00 = 1
                Exit Function
            Else
                Kill gsUpdPfad & "\" & gsUpdDatName
            End If
        End If
    Next i
  
Exit Function
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    ElseIf err.Number = 75 Then
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul1"
        Fehler.gsFunktion = "NEWfnCheck4UpdateDateiWKL00"
        Fehler.gsFehlertext = "Eventuell befinden sich schreibgeschützte Winkiss-Vorgängerversionen im Datenbank/IN Pfad"
        
        Fehlermeldung1
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul1"
        Fehler.gsFunktion = "NEWfnCheck4UpdateDateiWKL00"
        Fehler.gsFehlertext = "Bei der Überprüfung, ob ein Winkiss Update vorliegt, ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Function
Public Sub SichernBonDaten(cDruckZeile() As String, lAnzZeile As Long, kk_art As String, cKund As String, bPlusLeerZeilen As Boolean, Optional bBonNrgleichNull As Boolean = False)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim ctmp As String
    
    Dim lHeute As Long
    Dim lalt As Long
    Dim dSumme As Double
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cSumme As String
    Dim cUhrZeit As String
    Dim ierrz As Integer
    ierrz = 0
    '****************************************************************************************
    '* Hier kommt der neue Teil
    '****************************************************************************************
    If bBonNrgleichNull = True Then
        gdBonNr = 0
    End If
    lHeute = Fix(Now)
    
    ctmp = ""
    cUhrZeit = ""
    For lcount = 1 To lAnzZeile
        ctmp = ctmp & cDruckZeile(lcount)
        
        If InStr(cDruckZeile(lcount), Format$(lHeute, "DD.MM.YYYY")) > 0 Then
            cUhrZeit = cDruckZeile(lcount)
        End If
    Next lcount
    
    If bPlusLeerZeilen Then
    
        ctmp = ctmp & vbCrLf
        ctmp = ctmp & vbCrLf
        ctmp = ctmp & vbCrLf
        ctmp = ctmp & vbCrLf
        ctmp = ctmp & vbCrLf
        ctmp = ctmp & vbCrLf
        ctmp = ctmp & vbCrLf
        ctmp = ctmp & vbCrLf
        ctmp = ctmp & vbCrLf
    
    End If
    
    KonvertAsciiAnsi ctmp
    
    cUhrZeit = Format$(Now, "HH:MM:SS")
    
    If cUhrZeit < "00:00:00" Or cUhrZeit > "23:59:59" Then
        cUhrZeit = Format$(Now, "HH:MM:SS")
    End If
    
    cSQL = "Select * from KASSBON "
    cSQL = cSQL & "where DATUM = " & Trim$(Str$(lHeute)) & " "
    cSQL = cSQL & "and BONNR = " & Trim$(Str$(gdBonNr)) & " "
    cSQL = cSQL & "and KASNUM = " & gcKasNum & " "
    cSQL = cSQL & "and UHRZEIT = '" & cUhrZeit & "' "
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If rsrs.EOF Then
        rsrs.AddNew
        rsrs!Datum = lHeute
        rsrs!kasnum = Val(gcKasNum)
        rsrs!BONNR = gdBonNr
        rsrs!Uhrzeit = cUhrZeit
        rsrs!Betrag = gdSumme
        rsrs!BONTEXT = ctmp
        rsrs!FILIALE = Val(gcFilNr)
        rsrs!kk_art = kk_art
        rsrs!Kundnr = Val(cKund)
            
           
           
        If E_TSE_Aktiv And TSE_OK And gdSumme > 0 Then
        
            'TSE INFO SICHERN
            rsrs!TSESTART = R_StartTime
            rsrs!TSEEND = R_FinishTime
            rsrs!TSESERIAL = USB_serialNumber
            rsrs!TSETRANSACTION = R_TransactionNr
            rsrs!TSEClientID = E_ClientID
            rsrs!TSEFEHLER = TSE_Err
            rsrs!TSESTARTSIG = R_StartSignatur
            rsrs!TSEFINISHSIG = R_FinishSignatur
            rsrs!QRCODE = R_QRCodeAlsImgPath
            rsrs!TSEID = TSE_ID
            rsrs!STARTSIGZAHLER = R_START_SIG_Zaehler
            rsrs!FINISHSIGZAHLER = R_FINISH_SIG_Zaehler
         Else
            rsrs!TSEFEHLER = TSE_Err
        End If
        
        rsrs!SENDOK = False

        rsrs.Update
    Else
        rsrs.Edit
        rsrs!BONTEXT = rsrs!BONTEXT & ctmp
        
       If E_TSE_Aktiv And TSE_OK And gdSumme > 0 Then
        
            'TSE INFO SICHERN
            rsrs!TSESTART = R_StartTime
            rsrs!TSEEND = R_FinishTime
            rsrs!TSESERIAL = USB_serialNumber
            rsrs!TSETRANSACTION = R_TransactionNr
            rsrs!TSEClientID = E_ClientID
            rsrs!TSEFEHLER = TSE_Err
            rsrs!TSESTARTSIG = R_StartSignatur
            rsrs!TSEFINISHSIG = R_FinishSignatur
            rsrs!QRCODE = R_QRCodeAlsImgPath
            rsrs!TSEID = TSE_ID
            rsrs!STARTSIGZAHLER = R_START_SIG_Zaehler
            rsrs!FINISHSIGZAHLER = R_FINISH_SIG_Zaehler
         Else
            rsrs!TSEFEHLER = TSE_Err
        End If
        
        
        rsrs.Update
    End If
    rsrs.Close: Set rsrs = Nothing
     
    
    '*hier beginnt der neue Code für Kassbon.mdb
    Dim cPfad       As String
    Dim KASSBON_DB    As Database
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "GDPdU\KASSBON.MDB"
    
    Set KASSBON_DB = OpenDatabase(cPfad, False, False, "MS Access;PWD=" & gsKASSBON_Passwort)
    
    
    cSQL = "Select * from KASSBOND "
    cSQL = cSQL & "where DATUM = " & Trim$(Str$(lHeute)) & " "
    cSQL = cSQL & "and BONNR = " & Trim$(Str$(gdBonNr)) & " "
    cSQL = cSQL & "and KASNUM = " & gcKasNum & " "
    cSQL = cSQL & "and UHRZEIT = '" & cUhrZeit & "' "
    
    Set rsrs = KASSBON_DB.OpenRecordset(cSQL)
    If rsrs.EOF Then
    
        rsrs.AddNew
        rsrs!Datum = lHeute
        rsrs!kasnum = Val(gcKasNum)
        rsrs!BONNR = gdBonNr
        rsrs!Uhrzeit = cUhrZeit
        rsrs!Betrag = gdSumme
        rsrs!BONTEXT = ctmp
        rsrs!FILIALE = Val(gcFilNr)
        rsrs!kk_art = kk_art
        rsrs!Kundnr = Val(cKund)
        rsrs!SENDOK = False
        rsrs.Update
    Else
        rsrs.Edit
        rsrs!BONTEXT = rsrs!BONTEXT & ctmp
        rsrs.Update
    End If
    rsrs.Close: Set rsrs = Nothing
    
    KASSBON_DB.Close
    
    '*Ende ****hier beginnt der neue Code für Kassbon.mdb

    
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 3260 Then
        If ierrz < 5 Then
            ierrz = ierrz + 1
            Pause (1)
            Resume
        Else
            Fehler.gsDescr = err.Description
            Fehler.gsNumber = err.Number
            Fehler.gsFormular = "Modul1"
            Fehler.gsFunktion = "SichernBonDaten"
            Fehler.gsFehlertext = "Nach 5 sec ist ein Fehler aufgetreten."
            
            Fehlermeldung1
            Exit Sub
        End If
    Else
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul1"
        Fehler.gsFunktion = "SichernBonDaten"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub loesch(sTab As String)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW sTab, gdBase
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "loesch"
    Fehler.gsFehlertext = "Beim Löschen der Tabelle " & sTab & " ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Sub Hintergrundtabelle_kopieren(sTab1 As String, Optional sTab2 As String = "", Optional sTab3 As String = "", Optional sTab4 As String = "")
On Error GoTo LOKAL_ERROR

    Dim i               As Integer
    Dim sTabname        As String
    Dim slokalPfad      As String
    Dim sTabArray(3)    As String
    Dim sSQL            As String
    
    slokalPfad = gcDBPfad
    If Right(slokalPfad, 1) <> "\" Then
        slokalPfad = slokalPfad & "\"
    End If
    
    Dim dbWK As Database
    Set dbWK = OpenDatabase(slokalPfad & "kissdata.mdb", False, False, "MS Access;PWD=" & gsPasswort)
    
    sTabArray(0) = sTab1
    sTabArray(1) = sTab2
    sTabArray(2) = sTab3
    sTabArray(3) = sTab4
    
    For i = 0 To 3
        If sTabArray(i) = "" Then
            Exit For
        Else
            sTabname = sTabArray(i)
            
            loeschNEW sTabname, dbWK
        
            sSQL = "Select * "
            sSQL = sSQL & " into [;DATABASE=" & slokalPfad & "KISSDATA.MDB;pwd=" & gsPasswort & "]." & sTabname & " from " & sTabname & " "
            gdBase.Execute sSQL, dbFailOnError
        End If
    Next i
    
    dbWK.Close
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "Hintergrundtabelle_kopieren"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub dbPrintClose()
On Error GoTo LOKAL_ERROR

    dbPrintAusKissdata.Close
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "dbPrintClose"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub dbPrintOpen()
On Error GoTo LOKAL_ERROR

    Dim slokalPfad As String
    slokalPfad = gcDBPfad
    If Right(slokalPfad, 1) <> "\" Then
        slokalPfad = slokalPfad & "\"
    End If
    
    
    Set dbPrintAusKissdata = OpenDatabase(slokalPfad & "kissdata.mdb", False, False, "MS Access;PWD=" & gsPasswort)
    
    
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "dbPrintOpen"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub SQL_Befehl_ausführen(sSQL_Befehl As String)
    On Error GoTo LOKAL_ERROR
    
    
    
    If gbSQLSERVER = True Then
    
        If InStr(UCase(sSQL_Befehl), "DELETE ") > 0 Then
    
            gdBase.Execute sSQL_Befehl, dbSQLPassThrough + dbFailOnError
            
        ElseIf InStr(UCase(sSQL_Befehl), "DROP INDEX ") > 0 Then
            
            gdBase.Execute sSQL_Befehl, dbSQLPassThrough + dbFailOnError
            
        ElseIf InStr(UCase(sSQL_Befehl), "CREATE INDEX ") > 0 Then
            
            gdBase.Execute sSQL_Befehl, dbSQLPassThrough + dbFailOnError
            
        ElseIf InStr(UCase(sSQL_Befehl), "UPDATE ") > 0 Then
            
            gdBase.Execute sSQL_Befehl, dbSQLPassThrough + dbFailOnError
        
        ElseIf InStr(UCase(sSQL_Befehl), "CREATE TABLE ") > 0 Then
        
            sSQL_Befehl = SwapStr(UCase(sSQL_Befehl), " TEXT(", " varchar(")
            sSQL_Befehl = SwapStr(UCase(sSQL_Befehl), " LONG", " int")
            sSQL_Befehl = SwapStr(UCase(sSQL_Befehl), " DOUBLE", " float")
            sSQL_Befehl = SwapStr(UCase(sSQL_Befehl), " SINGLE", " real")
            sSQL_Befehl = SwapStr(UCase(sSQL_Befehl), " BYTE", " tinyint")
            sSQL_Befehl = SwapStr(UCase(sSQL_Befehl), " MEMO", " ntext")
            sSQL_Befehl = SwapStr(UCase(sSQL_Befehl), " INTEGER", " smallint")
            sSQL_Befehl = SwapStr(UCase(sSQL_Befehl), " AUTOINCREMENT", " Int IDENTITY")
            
            gdBase.Execute sSQL_Befehl, dbSQLPassThrough + dbFailOnError
            
        Else
            gdBase.Execute sSQL_Befehl, dbFailOnError
            
        End If

        
    Else
        gdBase.Execute sSQL_Befehl, dbFailOnError
    End If
    
    
    
    
     
    Exit Sub
LOKAL_ERROR:
    If err.Number = 3146 Then
        Resume Next
    Else
            Fehler.gsDescr = err.Description
            Fehler.gsNumber = err.Number
            Fehler.gsFormular = "Modul1"
            Fehler.gsFunktion = "SQL_Befehl_ausführen"
            Fehler.gsFehlertext = "Beim Ausführen eines SQL-Befehls ist ein Fehler aufgetreten." & sSQL_Befehl
            
            Fehlermeldung1
    End If
End Sub
Public Sub loeschNEW(sTab As String, db As Database)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    sSQL = "drop table " & sTab
    
    If gbSQLSERVER = True Then
        db.Execute sSQL, dbSQLPassThrough + dbFailOnError
    Else
    
        If db.name = gdBase.name Then
            db.Execute sSQL, dbFailOnError
        Else
            db.Execute sSQL, dbFailOnError
        End If
    
    End If
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 3376 Or err.Number = 3371 Or err.Number = 3167 Then
        Resume Next
    ElseIf err.Number = 3146 Then
        Resume Next
    ElseIf err.Number = 3262 Then
        sSQL = "Delete from " & sTab
    
        If gbSQLSERVER = True Then
            db.Execute sSQL, dbSQLPassThrough + dbFailOnError
        Else
            If db.name = gdBase.name Then
                schreibeProtokollDabaAblauf sSQL
                db.Execute sSQL, dbFailOnError
            Else
                db.Execute sSQL, dbFailOnError
            End If
        End If
        

        Exit Sub
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul1"
        Fehler.gsFunktion = "loeschNEW"
        Fehler.gsFehlertext = "Beim Löschen der Tabelle " & sTab & " in der Datenbank " & db.name & " ist ein Fehler aufgetreten."
        
        Fehlermeldung1
'        Resume Next
    End If
End Sub
Public Sub loeschapp(sTab As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    
    sSQL = "drop table " & sTab
    gdApp.Execute sSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 3376 Or err.Number = 3371 Then
        Resume Next

    Else
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul1"
        Fehler.gsFunktion = "loeschapp"
        Fehler.gsFehlertext = "Beim Löschen der Tabelle " & sTab & " ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Function fnHoleKundenNameMOD1(ctmp As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cName1 As String
    Dim cName2 As String
    
    If ctmp = "" Then
        Exit Function
    End If
    
    If IsNumeric(ctmp) = False Then
        Exit Function
    End If
    
    cSQL = "Select * from KUNDEN where KUNDNR = " & ctmp
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!name) Then
            cName2 = rsrs!name
        Else
            cName2 = ""
        End If
        
        If Not IsNull(rsrs!vorname) Then
            cName1 = rsrs!vorname
        Else
            cName1 = ""
        End If
        
        If cName1 = "" Then
            If Not IsNull(rsrs!titel) Then
                cName1 = rsrs!titel
            Else
                cName1 = ""
            End If
        End If
        
        If cName1 = "" Then
            If Not IsNull(rsrs!firma) Then
                cName1 = rsrs!firma
            Else
                cName1 = ""
            End If
        End If
            
        If cName1 = "" Then
            If Not IsNull(rsrs!geschlecht) Then
                cName1 = rsrs!geschlecht
                If cName1 = "F" Then
                    cName1 = "Firma"
                ElseIf cName1 = "W" Then
                    cName1 = "Frau"
                ElseIf cName1 = "M" Then
                    cName1 = "Herr"
                End If
            Else
                cName1 = ""
            End If
        End If
            
        If cName1 = "" Then
            cName1 = "Firma"
        End If
        
        fnHoleKundenNameMOD1 = cName1 & " " & cName2
        
    Else
        fnHoleKundenNameMOD1 = ""
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "fnHoleKundenNameMOD1"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub HoleNeueBonNrWKL20_NEU()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim rsRs2       As Recordset
    Dim lDatum      As Long
    
    lDatum = Fix(Now)

    cSQL = "Select * from AFCSTAT where ADATE = " & Trim$(Str$(lDatum)) & " and KASNUM = " & gcKasNum & " "
    Set rsRs2 = gdBase.OpenRecordset(cSQL)
    
    If Not rsRs2.EOF Then
        'Sicherheit
        
        cSQL = "Update AFCSTAT set BELEGNR = 999 where BELEGNR is null and ADATE = " & Trim$(Str$(lDatum)) & " and KASNUM = " & gcKasNum & " "
        gdBase.Execute cSQL, dbFailOnError
    
    
        cSQL = "Update AFCSTAT set BELEGNR = clng(BELEGNR) + 1 where ADATE = " & Trim$(Str$(lDatum)) & " and KASNUM = " & gcKasNum & " "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Select BELEGNR from AFCSTAT where ADATE = " & Trim$(Str$(lDatum)) & " and KASNUM = " & gcKasNum & " "
        Set rsrs = gdBase.OpenRecordset(cSQL)
        
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            
            If Not IsNull(rsrs!BELEGNR) Then
                gdBonNr = rsrs!BELEGNR
                
                If gbLokalModus = False Then
                    If gdBonNr < 1000 Then
                        gdBonNr = 1000
                    End If
                End If
            Else
                gdBonNr = 1000
            End If
            
            
        Else
            gdBonNr = 1000
        End If
        rsrs.Close: Set rsrs = Nothing
    Else
        gdBonNr = 1000
    End If
    rsRs2.Close: Set rsRs2 = Nothing
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "HoleNeueBonNrWKL20_NEU"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub HoleNeueBonNrWKL20()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim lDatum      As Long
    
    lDatum = Fix(Now)
    
    cSQL = "Select * from AFCSTAT where ADATE = " & Trim$(Str$(lDatum)) & " and KASNUM = " & gcKasNum & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!BELEGNR) Then
            gdBonNr = rsrs!BELEGNR
            gdBonNr = gdBonNr + 1
            
            If gbLokalModus = False Then
                If gdBonNr < 1000 Then
                    gdBonNr = 1000
                End If
            End If
        Else
            gdBonNr = 1000
        End If
    Else
        gdBonNr = 1000
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "HoleNeueBonNrWKL20"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub setThegdbaseNew()
    On Error GoTo LOKAL_ERROR

    gdBase.Close
    DabaPfadNew84
    
    
    
    If gcDBPfad <> "" Then
        If Not Modul6.FindFile(gcDBPfad, "kissdata.mdb") Then
            gcDBPfad = "C:\aLeer"
            Set gdBase = OpenDatabase(gcDBPfad & "\kissdata.mdb", False)
            Exit Sub
        End If
    Else
        gcDBPfad = "C:\aLeer"
        Set gdBase = OpenDatabase(gcDBPfad & "\kissdata.mdb", False)
        Exit Sub
    End If
'    Set gdBase = OpenDatabase(gcDBPfad & "\kissdata.mdb", False)
    Set gdBase = OpenDatabase(gcDBPfad & "\kissdata.mdb", False, False, "MS Access;PWD=" & gsPasswort)
    
    If NewTableSuchenDBKombi("ZZZ", gdBase) = False Then
        gcDBPfad = "C:\aLeer"
        Set gdBase = OpenDatabase(gcDBPfad & "\kissdata.mdb", False)
        Exit Sub
    End If
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "setThegdbaseNew"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub setThegdbaseLokal()
    On Error GoTo LOKAL_ERROR

    gdBase.Close
    gcDBPfad = "C:\aLeer"
    Set gdBase = OpenDatabase(gcDBPfad & "\kissdata.mdb", False)
               
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "setThegdbaseLokal"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub BerechneGrundPreis(dInhalt As Double, cInhaltBez As String, dVkPr As Double, cGrundInhalt As String, dGrundPreisDM As Double, dGrundPreisEur As Double)
    On Error GoTo LOKAL_ERROR
    
    '******************************************************
    '* vorhandene Maßeinheiten:
    '* L und ML -> Liter und Milliliter
    '* KG und G -> Kilogramm und Gramm
    '* M und CM -> Meter und Zentimeter
    '******************************************************
    
    '******************************************************
    '* dInhalt = Inhalt der Packung
    '* cInhaltBez = Maßeinheit, s.o.
    '* dVkPr = Verkaufspreis
    '* cGrundInh = Inhalt des Grundpreises + Maßeinheit
    '* dGrundPreisDM = Grundpreis in DM
    '* dGrundPreisEur = Grundpreis in EURO
    '******************************************************
    
    '********************************************************
    '* Umrechnungsmethoden:
    '*
    '* alles in L, KG und M wird auf Basis 1 heruntergerechnet
    '*
    '* alles in ML und G wird je nach dInhalt berechnet:
    '* 10 >= dInhalt >= 250 --> auf 100
    '* 250 > dInhalt --> 1000
    '*
    '* alles in CM wird je nach dInhalt berechnet:
    '* 1 >= dInhalt >= 25 --> auf 10
    '* 25 > dInhalt --> 100
    '*
    '* Wenn Rückgabewert -1 ist, dann ist kein Grundpreis erforderlich
    '* z.B. bei Kleinstmengen
    '*
    '********************************************************
    Select Case UCase(cInhaltBez)
        
        Case "L", "KG", "M", "WL"
'            If dInhalt < 1 Then
'                cGrundInhalt = ""
'                dGrundPreisDM = -1
'                dGrundPreisEur = -1
'            End If
            'Berechnung des Grundpreises auf Inhalt 1 L / KG / M
'            If dInhalt < 1 Then
'                cGrundInhalt = "1 " & cInhaltbez
'                dGrundPreisDM = (1 / dInhalt) * dVkPr
'                dGrundPreisEur = dGrundPreisDM
'            End If
            
            If dInhalt > 0 And dInhalt <= 1000 Then
                cGrundInhalt = "1 " & cInhaltBez
                dGrundPreisDM = (1 / dInhalt) * dVkPr
                dGrundPreisEur = dGrundPreisDM
            End If
        
        Case "ML", "G"
            If dInhalt < 1 Then
                cGrundInhalt = ""
                dGrundPreisDM = -1
                dGrundPreisEur = -1
            End If
            
            If dInhalt >= 1 And dInhalt < 10 Then
                cGrundInhalt = "10 " & cInhaltBez
                dGrundPreisDM = (10 / dInhalt) * dVkPr
                dGrundPreisEur = dGrundPreisDM
            End If
            
            
            If dInhalt >= 10 And dInhalt <= 250 Then
                cGrundInhalt = "100 " & cInhaltBez
                dGrundPreisDM = (100 / dInhalt) * dVkPr
                dGrundPreisEur = dGrundPreisDM
            End If
            
            If dInhalt > 250 Then
                cGrundInhalt = "1000 " & cInhaltBez
                dGrundPreisDM = (1000 / dInhalt) * dVkPr
                dGrundPreisEur = dGrundPreisDM
            End If
            
        Case "CM"
            If dInhalt < 1 Then
                cGrundInhalt = ""
                dGrundPreisDM = -1
                dGrundPreisEur = -1
            End If
            
            If dInhalt >= 1 And dInhalt <= 25 Then
                cGrundInhalt = "10 " & cInhaltBez
                dGrundPreisDM = (10 / dInhalt) * dVkPr
                dGrundPreisEur = dGrundPreisDM
            End If
            
            If dInhalt > 25 Then
                cGrundInhalt = "100 " & cInhaltBez
                dGrundPreisDM = (100 / dInhalt) * dVkPr
                dGrundPreisEur = dGrundPreisDM
            End If
            
        Case "ST"
            If dInhalt < 1 Then
                cGrundInhalt = ""
                dGrundPreisDM = -1
                dGrundPreisEur = -1
            End If
            
            If dInhalt >= 1 And dInhalt <= 25 Then
                cGrundInhalt = "1 " & cInhaltBez
                dGrundPreisDM = (1 / dInhalt) * dVkPr
                dGrundPreisEur = dGrundPreisDM
            End If
            
            If dInhalt > 25 Then
                cGrundInhalt = "10 " & cInhaltBez
                dGrundPreisDM = (10 / dInhalt) * dVkPr
                dGrundPreisEur = dGrundPreisDM
            End If
        
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "BerechneGrundPreis"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Function ean13(sCode As String) As String
On Error GoTo Fehler
    Dim iZähler As Integer
    Dim rsCode As DAO.Recordset
    Dim iAnzahl As Integer
    Dim sCodeEinzel As String
    Dim sCodeGesamt As String
    Dim sCodeNeu As String
    
    'Prüfung ob Zahl in Ordnung ist
    'Prüfziffer abfragen = 13
    If Not Mod10(sCode) Then ean13 = "#Fehler": Exit Function
    If Not IsNumeric(sCode) Then ean13 = "#Text": Exit Function 'Nur Zahlen
    'Tabelle mit der Zuweisung öffnen
    Set rsCode = gdBase.OpenRecordset("tabCode", dbOpenDynaset)
    '*******************
    '1.Zahl plus *
     iZähler = 1
    Do While iZähler <= 13
        sCodeEinzel = Right(Left(sCode, iZähler), 1) 'Immer die nächste Zahl
        If iZähler = 1 Then
            sCodeEinzel = Left(sCode, 1) & " *"
            sCodeGesamt = sCodeEinzel
        Else
            If iZähler >= 8 Then
                rsCode.FindFirst "Stelle_Ziffer1 = 8"
              Else                                             'NEU eingefügt!
                If iZähler = 2 Then                            'NEU eingefügt!
                    rsCode.FindFirst "Stelle_Ziffer1 = 2"      'NEU eingefügt!
                  Else
                    'in der Tabelle die Stelle suchen GEÄNDERT!!
                    rsCode.FindFirst "Stelle_Ziffer1 = " & iZähler & _
                                     Left(sCode, 1)
                End If                                         'NEU eingefügt!
            End If
            'Richtige spalte auslesen
            sCodeEinzel = rsCode.Fields("A_" & sCodeEinzel)
            sCodeGesamt = sCodeGesamt & sCodeEinzel 'Code zusammensetzen
'            Debug.Print " Stelle: " & iZähler & " " & sCodeEinzel
        End If
        iZähler = iZähler + 1
    Loop
    'An 10. Stelle ein # am ende ein * und ein
    sCodeGesamt = Left(sCodeGesamt, 9) & "#" & Right(sCodeGesamt, 6)
    ean13 = sCodeGesamt & "* "
    
Fehler_Exit:
    rsCode.Close
    Set rsCode = Nothing
    Exit Function
Fehler:
    MsgBox "Folgender Fehler ist aufgetaucht: " & vbCrLf & _
            err.Number & ": " & vbCrLf & _
            err.Description, vbCritical, "Fehler..."
    ean13 = "#Error"
    Resume Fehler_Exit
End Function
Public Function fncodiereEan13$(chaine$)
  'Cette fonction est regie par la Licence Generale Publique Amoindrie GNU (GNU LGPL)
  'This function is governed by the GNU Lesser General Public License (GNU LGPL)
  'V 1.1.1
  'Parametres : une chaine de 12 chiffres
  'Parameters : a 12 digits length string
  'Retour : * une chaine qui, affichee avec la police EAN13.TTF, donne le code barre
  '         * une chaine vide si parametre fourni incorrect
  'Return : * a string which give the bar code when it is dispayed with EAN13.TTF font
  '         * an empty string if the supplied parameter is no good
  Dim i%, checksum%, first%, CodeBarre$, tableA As Boolean
  fncodiereEan13$ = ""
  'Verifier qu'il y a 12 caracteres
  'Check for 12 characters
  If Len(chaine$) = 12 Then
    'Et que ce sont bien des chiffres
    'And they are really digits
    For i% = 1 To 12
      If Asc(Mid$(chaine$, i%, 1)) < 48 Or Asc(Mid$(chaine$, i%, 1)) > 57 Then
        i% = 0
        Exit For
      End If
    Next
    If i% = 13 Then
      'Calcul de la cle de controle
      'Calculation of the checksum
      For i% = 12 To 1 Step -2
        checksum% = checksum% + Val(Mid$(chaine$, i%, 1))
      Next
      checksum% = checksum% * 3
      For i% = 11 To 1 Step -2
        checksum% = checksum% + Val(Mid$(chaine$, i%, 1))
      Next
      chaine$ = chaine$ & (10 - checksum% Mod 10) Mod 10
      'Le premier chiffre est pris tel quel, le deuxieme vient de la table A
      'The first digit is taken just as it is, the second one come from table A
      CodeBarre$ = Left$(chaine$, 1) & Chr$(65 + Val(Mid$(chaine$, 2, 1)))
      first% = Val(Left$(chaine$, 1))
      For i% = 3 To 7
        tableA = False
         Select Case i%
         Case 3
           Select Case first%
           Case 0 To 3
             tableA = True
           End Select
         Case 4
           Select Case first%
           Case 0, 4, 7, 8
             tableA = True
           End Select
         Case 5
           Select Case first%
           Case 0, 1, 4, 5, 9
             tableA = True
           End Select
         Case 6
           Select Case first%
           Case 0, 2, 5, 6, 7
             tableA = True
           End Select
         Case 7
           Select Case first%
           Case 0, 3, 6, 8, 9
             tableA = True
           End Select
         End Select
       If tableA Then
         CodeBarre$ = CodeBarre$ & Chr$(65 + Val(Mid$(chaine$, i%, 1)))
       Else
         CodeBarre$ = CodeBarre$ & Chr$(75 + Val(Mid$(chaine$, i%, 1)))
       End If
     Next
      CodeBarre$ = CodeBarre$ & "*"   'Ajout separateur central / Add middle separator
      For i% = 8 To 13
        CodeBarre$ = CodeBarre$ & Chr$(97 + Val(Mid$(chaine$, i%, 1)))
      Next
      CodeBarre$ = CodeBarre$ & "+"   'Ajout de la marque de fin / Add end mark
      fncodiereEan13$ = CodeBarre$
    End If
  End If
End Function

Function Mod10(sCode) As Boolean
    Dim i As Integer
    Dim bZahl As Integer
    Dim iSumme As Integer

    Mod10 = True
    If Len(sCode) <> 13 Then
        MsgBox "Code ist zu klein"
        Exit Function
    End If
    For i = 1 To Len(sCode) - 1
        bZahl = Right(Left(sCode, i), 1)
        If i Mod 2 = 0 Then bZahl = bZahl * 3
        iSumme = iSumme + bZahl
    Next i
    iSumme = 10 - (iSumme Mod 10)
    
    If (iSumme = Right(sCode, 1)) Or (Right(iSumme, 1) = Right(sCode, 1)) Then
        Mod10 = True
    Else
        Mod10 = False
    End If
End Function
Public Function fnCodiereEANCode(cEAN As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim cTeil1(0 To 9) As String
    Dim cTeil2(0 To 9) As String
    Dim cTeilA(0 To 9) As String
    Dim cTeilB(0 To 9) As String
    Dim cTeilC(0 To 9) As String
    Dim cZeichen As String
    Dim lcount As Long
    Dim cZiel As String
    
    fnCodiereEANCode = ""
    
    cTeil1(0) = "p" 'A
    cTeil1(1) = "q"
    cTeil1(2) = "w"
    cTeil1(3) = "e"
    cTeil1(4) = "r"
    cTeil1(5) = "t"
    cTeil1(6) = "z"
    cTeil1(7) = "u"
    cTeil1(8) = "i"
    cTeil1(9) = "o"
    
    cTeil2(0) = "-" 'C
    cTeil2(1) = "y"
    cTeil2(2) = "x"
    cTeil2(3) = "c"
    cTeil2(4) = "v"
    cTeil2(5) = "b"
    cTeil2(6) = "n"
    cTeil2(7) = "m"
    cTeil2(8) = ","
    cTeil2(9) = "."
    
'    cTeilB(0) = "p" 'B
'    cTeilB(1) = "a"
'    cTeilB(2) = "s"
'    cTeilB(3) = "d"
'    cTeilB(4) = "f"
'    cTeilB(5) = "g"
'    cTeilB(6) = "h"
'    cTeilB(7) = "j"
'    cTeilB(8) = "k"
'    cTeilB(9) = "l"

    If Len(cEAN) = 8 Then
        
        cZiel = Chr$(42)
        
        For lcount = 1 To 4
            cZeichen = Mid(cEAN, lcount, 1)
            cZiel = cZiel & cTeil1(Val(cZeichen))
        Next lcount
        
        cZiel = cZiel & "#"
        
        For lcount = 5 To 8
            cZeichen = Mid(cEAN, lcount, 1)
            cZiel = cZiel & cTeil2(Val(cZeichen))
        Next lcount
        
        cZiel = cZiel & Chr$(42)
    
    End If
    
    
    
    
    
    
    cTeilA(0) = "ö" 'A
    cTeilA(1) = "q"
    cTeilA(2) = "w"
    cTeilA(3) = "e"
    cTeilA(4) = "r"
    cTeilA(5) = "t"
    cTeilA(6) = "z"
    cTeilA(7) = "u"
    cTeilA(8) = "i"
    cTeilA(9) = "o"
    
    
    
    cTeilB(0) = "p" 'B
    cTeilB(1) = "a"
    cTeilB(2) = "s"
    cTeilB(3) = "d"
    cTeilB(4) = "f"
    cTeilB(5) = "g"
    cTeilB(6) = "z"
    cTeilB(7) = "j"
    cTeilB(8) = "k"
    cTeilB(9) = "l"
    
    cTeilC(0) = "-" 'C
    cTeilC(1) = "y"
    cTeilC(2) = "x"
    cTeilC(3) = "c"
    cTeilC(4) = "v"
    cTeilC(5) = "b"
    cTeilC(6) = "n"
    cTeilC(7) = "m"
    cTeilC(8) = ","
    cTeilC(9) = "."
    
    
    
    
    
    
    
    If Len(cEAN) = 13 Then
        Dim cpruef As String
        cpruef = Right(cEAN, 1)
        
        cZiel = Left(cEAN, 1)
        cZiel = cZiel & Chr$(42)
        
        Select Case cpruef
        
        Case "0"
        
            cZeichen = Mid(cEAN, 2, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
            cZeichen = Mid(cEAN, 3, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
            cZeichen = Mid(cEAN, 4, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
            cZeichen = Mid(cEAN, 5, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
            cZeichen = Mid(cEAN, 6, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
            cZeichen = Mid(cEAN, 7, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
        
        
        Case "1"
        
            cZeichen = Mid(cEAN, 2, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
            cZeichen = Mid(cEAN, 3, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
            cZeichen = Mid(cEAN, 4, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
            cZeichen = Mid(cEAN, 5, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
            cZeichen = Mid(cEAN, 6, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
            cZeichen = Mid(cEAN, 7, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
            
        Case "2"
        
            cZeichen = Mid(cEAN, 2, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
            cZeichen = Mid(cEAN, 3, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
            cZeichen = Mid(cEAN, 4, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
            cZeichen = Mid(cEAN, 5, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
            cZeichen = Mid(cEAN, 6, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
            cZeichen = Mid(cEAN, 7, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
        

        Case "3"
            
            cZeichen = Mid(cEAN, 2, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
            cZeichen = Mid(cEAN, 3, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
            cZeichen = Mid(cEAN, 4, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
            cZeichen = Mid(cEAN, 5, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
            cZeichen = Mid(cEAN, 6, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
            cZeichen = Mid(cEAN, 7, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
            
        Case "4"
            
            cZeichen = Mid(cEAN, 2, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
            cZeichen = Mid(cEAN, 3, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
            cZeichen = Mid(cEAN, 4, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
            cZeichen = Mid(cEAN, 5, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
            cZeichen = Mid(cEAN, 6, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
            cZeichen = Mid(cEAN, 7, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
            
        Case "5"
            
            cZeichen = Mid(cEAN, 2, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
            cZeichen = Mid(cEAN, 3, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
            cZeichen = Mid(cEAN, 4, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
            cZeichen = Mid(cEAN, 5, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
            cZeichen = Mid(cEAN, 6, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
            cZeichen = Mid(cEAN, 7, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
        
        Case "6"
            
            cZeichen = Mid(cEAN, 2, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
            cZeichen = Mid(cEAN, 3, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
            cZeichen = Mid(cEAN, 4, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
            cZeichen = Mid(cEAN, 5, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
            cZeichen = Mid(cEAN, 6, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
            cZeichen = Mid(cEAN, 7, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
            
        Case "7"
            
            cZeichen = Mid(cEAN, 2, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
            cZeichen = Mid(cEAN, 3, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
            cZeichen = Mid(cEAN, 4, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
            cZeichen = Mid(cEAN, 5, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
            cZeichen = Mid(cEAN, 6, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
            cZeichen = Mid(cEAN, 7, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
        
        Case "8"
            
            cZeichen = Mid(cEAN, 2, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
            cZeichen = Mid(cEAN, 3, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
            cZeichen = Mid(cEAN, 4, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
            cZeichen = Mid(cEAN, 5, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
            cZeichen = Mid(cEAN, 6, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
            cZeichen = Mid(cEAN, 7, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
        
        Case "9"
        
            cZeichen = Mid(cEAN, 2, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
            cZeichen = Mid(cEAN, 3, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
            cZeichen = Mid(cEAN, 4, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
            cZeichen = Mid(cEAN, 5, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
            cZeichen = Mid(cEAN, 6, 1): cZiel = cZiel & cTeilB(Val(cZeichen))
            cZeichen = Mid(cEAN, 7, 1): cZiel = cZiel & cTeilA(Val(cZeichen))
        
        
        
        End Select
       
        
        cZiel = cZiel & "#"
        
        For lcount = 8 To 13
            cZeichen = Mid(cEAN, lcount, 1)
            cZiel = cZiel & cTeilC(Val(cZeichen))
        Next lcount
        
        cZiel = cZiel & Chr$(42)
    
    End If
    
    fnCodiereEANCode = cZiel
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "fnCodiereEANCode"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function fnCodiereEAN13Code(cEAN As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim cTeil1(0 To 9) As String
    Dim cTeil2(0 To 9) As String
    Dim cTeilA(0 To 9) As String
    Dim cTeilB(0 To 9) As String
    Dim cTeilC(0 To 9) As String
    Dim cZeichen As String
    Dim lcount As Long
    Dim cZiel As String
    
    fnCodiereEAN13Code = ""
    
    cTeil1(0) = "p" 'A
    cTeil1(1) = "q"
    cTeil1(2) = "w"
    cTeil1(3) = "e"
    cTeil1(4) = "r"
    cTeil1(5) = "t"
    cTeil1(6) = "z"
    cTeil1(7) = "u"
    cTeil1(8) = "i"
    cTeil1(9) = "o"
    
    cTeil2(0) = "-" 'C
    cTeil2(1) = "y"
    cTeil2(2) = "x"
    cTeil2(3) = "c"
    cTeil2(4) = "v"
    cTeil2(5) = "b"
    cTeil2(6) = "n"
    cTeil2(7) = "m"
    cTeil2(8) = ","
    cTeil2(9) = "."
    
'    cTeilB(0) = "p" 'B
'    cTeilB(1) = "a"
'    cTeilB(2) = "s"
'    cTeilB(3) = "d"
'    cTeilB(4) = "f"
'    cTeilB(5) = "g"
'    cTeilB(6) = "h"
'    cTeilB(7) = "j"
'    cTeilB(8) = "k"
'    cTeilB(9) = "l"

    If Len(cEAN) = 13 Then
        
        cZiel = Chr$(42)
        
        For lcount = 1 To 4
            cZeichen = Mid(cEAN, lcount, 1)
            cZiel = cZiel & cTeil1(Val(cZeichen))
        Next lcount
        
        cZiel = cZiel & "#"
        
        For lcount = 5 To 8
            cZeichen = Mid(cEAN, lcount, 1)
            cZiel = cZiel & cTeil2(Val(cZeichen))
        Next lcount
        
        cZiel = cZiel & Chr$(42)
    
    End If
    
    
    
    
    
    
    cTeilA(0) = "ö" 'A
    cTeilA(1) = "q"
    cTeilA(2) = "w"
    cTeilA(3) = "e"
    cTeilA(4) = "r"
    cTeilA(5) = "t"
    cTeilA(6) = "z"
    cTeilA(7) = "u"
    cTeilA(8) = "i"
    cTeilA(9) = "o"
    
    
    
    cTeilB(0) = "p" 'B
    cTeilB(1) = "a"
    cTeilB(2) = "s"
    cTeilB(3) = "d"
    cTeilB(4) = "f"
    cTeilB(5) = "g"
    cTeilB(6) = "z"
    cTeilB(7) = "j"
    cTeilB(8) = "k"
    cTeilB(9) = "l"
    
    cTeilC(0) = "-" 'C
    cTeilC(1) = "y"
    cTeilC(2) = "x"
    cTeilC(3) = "c"
    cTeilC(4) = "v"
    cTeilC(5) = "b"
    cTeilC(6) = "n"
    cTeilC(7) = "m"
    cTeilC(8) = ","
    cTeilC(9) = "."
    
    
    '1. Zeichen
    cZiel = Left(cEAN, 1)
    
    '2.Zeichen
    cZiel = cZiel & " "
    
    '3.Zeichen
    cZiel = cZiel & "*"
    
    '4.Zeichen
    Select Case Mid(cEAN, 2, 1)
        Case 1: cZiel = cZiel & "q"
        Case 2: cZiel = cZiel & "w"
        Case 3: cZiel = cZiel & "e"
        Case 4: cZiel = cZiel & "r"
        Case 5: cZiel = cZiel & "t"
        Case 6: cZiel = cZiel & "z"
        Case 7: cZiel = cZiel & "u"
        Case 8: cZiel = cZiel & "i"
        Case 9: cZiel = cZiel & "o"
        Case 0: cZiel = cZiel & "p"
    End Select
    
    '5.Zeichen
    Select Case Mid(cEAN, 3, 1)
        Case 1: cZiel = cZiel & "q"
        Case 2: cZiel = cZiel & "w"
        Case 3: cZiel = cZiel & "e"
        Case 4: cZiel = cZiel & "f"
        Case 5: cZiel = cZiel & "g"
        Case 6: cZiel = cZiel & "h"
        Case 7: cZiel = cZiel & "j"
        Case 8: cZiel = cZiel & "k"
        Case 9: cZiel = cZiel & "l"
        Case 0: cZiel = cZiel & "p"
    End Select
    
    '6.Zeichen
    Select Case Mid(cEAN, 4, 1)
        Case 1: cZiel = cZiel & "a"
        Case 2: cZiel = cZiel & "s"
        Case 3: cZiel = cZiel & "d"
        Case 4: cZiel = cZiel & "r"
        Case 5: cZiel = cZiel & "g"
        Case 6: cZiel = cZiel & "h"
        Case 7: cZiel = cZiel & "u"
        Case 8: cZiel = cZiel & "i"
        Case 9: cZiel = cZiel & "l"
        Case 0: cZiel = cZiel & "p"
    End Select
    
    '7.Zeichen 5.Auswerten
    Select Case Mid(cEAN, 5, 1)
        Case 1: cZiel = cZiel & "q"
        Case 2: cZiel = cZiel & "s"
        Case 3: cZiel = cZiel & "e" 'd
        Case 4: cZiel = cZiel & "r"
        Case 5: cZiel = cZiel & "t"
        Case 6: cZiel = cZiel & "h"
        Case 7: cZiel = cZiel & "j"
        Case 8: cZiel = cZiel & "k"
        Case 9: cZiel = cZiel & "o"
        Case 0: cZiel = cZiel & "p"
    End Select
    
    '8.Zeichen 6.Auswerten
    Select Case Mid(cEAN, 6, 1)
        Case 1: cZiel = cZiel & "a"
        Case 2: cZiel = cZiel & "w"
        Case 3: cZiel = cZiel & "d"
        Case 4: cZiel = cZiel & "f"
        Case 5: cZiel = cZiel & "t"
        Case 6: cZiel = cZiel & "z"
        Case 7: cZiel = cZiel & "u"
        Case 8: cZiel = cZiel & "k"
        Case 9: cZiel = cZiel & "l"
        Case 0: cZiel = cZiel & "p"
    End Select
    
    '9.Zeichen 7.Auswerten
    Select Case Mid(cEAN, 7, 1)
        Case 1: cZiel = cZiel & "a"
        Case 2: cZiel = cZiel & "s"
        Case 3: cZiel = cZiel & "e"
        Case 4: cZiel = cZiel & "f"
        Case 5: cZiel = cZiel & "g"
        Case 6: cZiel = cZiel & "z"
        Case 7: cZiel = cZiel & "j"
        Case 8: cZiel = cZiel & "i"
        Case 9: cZiel = cZiel & "o"
        Case 0: cZiel = cZiel & "p"
    End Select
    
  
    '10.Zeichen immer #
    cZiel = cZiel & "#"
    
    '11.Zeichen 8.Auswerten
    Select Case Mid(cEAN, 8, 1)
        Case 1: cZiel = cZiel & "y"
        Case 2: cZiel = cZiel & "x"
        Case 3: cZiel = cZiel & "c"
        Case 4: cZiel = cZiel & "v"
        Case 5: cZiel = cZiel & "b"
        Case 6: cZiel = cZiel & "n"
        Case 7: cZiel = cZiel & "m"
        Case 8: cZiel = cZiel & ","
        Case 9: cZiel = cZiel & "."
        Case 0: cZiel = cZiel & "-"
    End Select
    
    '12.Zeichen 9.Auswerten
    Select Case Mid(cEAN, 9, 1)
        Case 1: cZiel = cZiel & "y"
        Case 2: cZiel = cZiel & "x"
        Case 3: cZiel = cZiel & "c"
        Case 4: cZiel = cZiel & "v"
        Case 5: cZiel = cZiel & "b"
        Case 6: cZiel = cZiel & "n"
        Case 7: cZiel = cZiel & "m"
        Case 8: cZiel = cZiel & ","
        Case 9: cZiel = cZiel & "."
        Case 0: cZiel = cZiel & "-"
    End Select
    
    '13.Zeichen 10.Auswerten
    Select Case Mid(cEAN, 10, 1)
        Case 1: cZiel = cZiel & "y"
        Case 2: cZiel = cZiel & "x"
        Case 3: cZiel = cZiel & "c"
        Case 4: cZiel = cZiel & "v"
        Case 5: cZiel = cZiel & "b"
        Case 6: cZiel = cZiel & "n"
        Case 7: cZiel = cZiel & "m"
        Case 8: cZiel = cZiel & ","
        Case 9: cZiel = cZiel & "."
        Case 0: cZiel = cZiel & "-"
    End Select
   
    '14.Zeichen 11.Auswerten
    Select Case Mid(cEAN, 11, 1)
        Case 1: cZiel = cZiel & "y"
        Case 2: cZiel = cZiel & "x"
        Case 3: cZiel = cZiel & "c"
        Case 4: cZiel = cZiel & "v"
        Case 5: cZiel = cZiel & "b"
        Case 6: cZiel = cZiel & "n"
        Case 7: cZiel = cZiel & "m"
        Case 8: cZiel = cZiel & ","
        Case 9: cZiel = cZiel & "."
        Case 0: cZiel = cZiel & "-"
    End Select
    
    '15.Zeichen 12.Auswerten
    Select Case Mid(cEAN, 12, 1)
        Case 1: cZiel = cZiel & "y"
        Case 2: cZiel = cZiel & "x"
        Case 3: cZiel = cZiel & "c"
        Case 4: cZiel = cZiel & "v"
        Case 5: cZiel = cZiel & "b"
        Case 6: cZiel = cZiel & "n"
        Case 7: cZiel = cZiel & "m"
        Case 8: cZiel = cZiel & ","
        Case 9: cZiel = cZiel & "."
        Case 0: cZiel = cZiel & "-"
    End Select
    
    '16.Zeichen 13.Auswerten
    Select Case Mid(cEAN, 13, 1)
        Case 1: cZiel = cZiel & "y"
        Case 2: cZiel = cZiel & "x"
        Case 3: cZiel = cZiel & "c"
        Case 4: cZiel = cZiel & "v"
        Case 5: cZiel = cZiel & "b"
        Case 6: cZiel = cZiel & "n"
        Case 7: cZiel = cZiel & "m"
        Case 8: cZiel = cZiel & ","
        Case 9: cZiel = cZiel & "."
        Case 0: cZiel = cZiel & "-"
    End Select
    
   

    '17. Zeichen:    “*“ -- IMMER --
    cZiel = cZiel & "*"
    '18. Zeichen:    “ “ (Leertaste) -- IMMER --
    cZiel = cZiel & " "

    
    
    
    
    
    
    
    
    
    
    
 
    
    fnCodiereEAN13Code = cZiel
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "fnCodiereEAN13Code"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function fnEntferneLeerzeichen(cText As String) As String
    Dim cLetzter As String
    Dim cZiel As String
    Dim lPos As Long
    Dim cZeichen As String
    
    'Überflüssige Leerzeichen herausfiltern
    cLetzter = ""
    cZiel = ""
    For lPos = 1 To Len(cText)
        cZeichen = Mid(cText, lPos, 1)
        If cZeichen = " " Then
            If cZeichen = cLetzter Then
                'unterdrücken
            Else
                cZiel = cZiel & cZeichen
            End If
            cLetzter = cZeichen
        Else
            cZiel = cZiel & cZeichen
            cLetzter = cZeichen
        End If
    Next lPos
    fnEntferneLeerzeichen = Trim$(cZiel)
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "fnEntferneLeerzeichen"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function fnMoveArtNr2EAN8_begin980(cArtNr As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim lcount As Long
    Dim lPruefZiffer As Long
    
    If Len(cArtNr) > 6 Then 'hier
        fnMoveArtNr2EAN8_begin980 = ""
        Exit Function
    End If
    
    If Len(cArtNr) < 4 Then 'hier
        cArtNr = String$(4 - Len(cArtNr), "0") & cArtNr 'hier
    End If
    
    cArtNr = "980" & cArtNr 'hier
    lPruefZiffer = 0
    For lcount = 1 To 7 'hier
        cZeichen = Mid(cArtNr, lcount, 1)
        If lcount / 2 = Int(lcount / 2) Then
            lPruefZiffer = lPruefZiffer + Val(cZeichen)
        Else
            lPruefZiffer = lPruefZiffer + (Val(cZeichen) * 3)
        End If
    Next lcount
    lPruefZiffer = lPruefZiffer Mod 10
    If lPruefZiffer > 0 Then
        lPruefZiffer = 10 - lPruefZiffer
    End If
    
    cArtNr = cArtNr & Trim$(Str$(lPruefZiffer))
    
    fnMoveArtNr2EAN8_begin980 = cArtNr
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "fnMoveArtNr2EAN8_begin980"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function fnMoveKundnr2EAN8(cArtNr As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim lcount As Long
    Dim lPruefZiffer As Long
    
    lPruefZiffer = 0
    For lcount = 1 To 7 'hier
        cZeichen = Mid(cArtNr, lcount, 1)
        If lcount / 2 = Int(lcount / 2) Then
            lPruefZiffer = lPruefZiffer + Val(cZeichen)
        Else
            lPruefZiffer = lPruefZiffer + (Val(cZeichen) * 3)
        End If
    Next lcount
    lPruefZiffer = lPruefZiffer Mod 10
    If lPruefZiffer > 0 Then
        lPruefZiffer = 10 - lPruefZiffer
    End If
    
    cArtNr = cArtNr & Trim$(Str$(lPruefZiffer))
    
    fnMoveKundnr2EAN8 = cArtNr
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "fnMoveKundnr2EAN8"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function fnMoveArtNr2EAN8(cArtNr As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim lcount As Long
    Dim lPruefZiffer As Long
    
    If Len(cArtNr) > 6 Then 'hier
        fnMoveArtNr2EAN8 = ""
        Exit Function
    End If
    
    If Len(cArtNr) < 6 Then 'hier
        cArtNr = String$(6 - Len(cArtNr), "0") & cArtNr 'hier
    End If
    
    cArtNr = "2" & cArtNr 'hier
    lPruefZiffer = 0
    For lcount = 1 To 7 'hier
        cZeichen = Mid(cArtNr, lcount, 1)
        If lcount / 2 = Int(lcount / 2) Then
            lPruefZiffer = lPruefZiffer + Val(cZeichen)
        Else
            lPruefZiffer = lPruefZiffer + (Val(cZeichen) * 3)
        End If
    Next lcount
    lPruefZiffer = lPruefZiffer Mod 10
    If lPruefZiffer > 0 Then
        lPruefZiffer = 10 - lPruefZiffer
    End If
    
    cArtNr = cArtNr & Trim$(Str$(lPruefZiffer))
    
    fnMoveArtNr2EAN8 = cArtNr
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "fnMoveArtNr2EAN8"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function

Public Function fnMoveGutschnr2EAN13(cGutschnr As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim lcount As Long
    Dim lPruefZiffer As Long
    Dim lPos As Long
    Dim lSumme As Long
    Dim lWert As Long
    
    If Len(cGutschnr) > 12 Then 'hier
        fnMoveGutschnr2EAN13 = ""
        Exit Function
    End If
    
    If Len(cGutschnr) < 12 Then 'hier
        cGutschnr = String$(12 - Len(cGutschnr), "0") & cGutschnr
    End If
    
    lPruefZiffer = 0
    
    For lPos = 1 To 12
        cZeichen = Mid(cGutschnr, lPos, 1)
        If lPos / 2 = Int(lPos / 2) Then
            lSumme = lSumme + (Val(cZeichen) * 3)
        Else
            lSumme = lSumme + Val(cZeichen)
        End If
    Next lPos
    lWert = lSumme Mod 10
    If lWert > 0 Then
        lWert = 10 - lWert
    End If
    
    fnMoveGutschnr2EAN13 = cGutschnr & CStr(lWert)
''    MsgBox Len(fnMoveGutschnr2EAN13)
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "fnMoveGutschnr2EAN13"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function fnMoveNr2EAN8(cArtNr As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim lcount As Long
    Dim lPruefZiffer As Long
    
    If Len(cArtNr) > 7 Then 'hier
        fnMoveNr2EAN8 = ""
        Exit Function
    End If
    
    If Len(cArtNr) < 7 Then 'hier
        cArtNr = String$(7 - Len(cArtNr), "0") & cArtNr 'hier
    End If
    
    cArtNr = cArtNr 'hier
    lPruefZiffer = 0
    For lcount = 1 To 7 'hier
        cZeichen = Mid(cArtNr, lcount, 1)
        If lcount / 2 = Int(lcount / 2) Then
            lPruefZiffer = lPruefZiffer + Val(cZeichen)
        Else
            lPruefZiffer = lPruefZiffer + (Val(cZeichen) * 3)
        End If
    Next lcount
    lPruefZiffer = lPruefZiffer Mod 10
    If lPruefZiffer > 0 Then
        lPruefZiffer = 10 - lPruefZiffer
    End If
    
    cArtNr = cArtNr & Trim$(Str$(lPruefZiffer))
    
    fnMoveNr2EAN8 = cArtNr
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "fnMoveNr2EAN8"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function fnEncrypt(cFeld) As String
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim cZeichen As String
    Dim lWert As Long
    Dim lWertMerker As Long
    Dim lZiel As Long
    Dim cZiel As String
    
    fnEncrypt = ""
    
    lWertMerker = 58
    For lcount = 1 To Len(cFeld)
        cZeichen = Mid(cFeld, lcount, 1)
        lWert = Asc(cZeichen)
        lZiel = lWertMerker + lWert
        If lZiel > 255 Then
            lZiel = lZiel - 255
        End If
        cZeichen = Chr$(lZiel)
        cZiel = cZiel & cZeichen
        lWertMerker = lWert
    Next lcount
    
    'MsgBox cZiel
    
    fnEncrypt = cZiel
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "fnEncrypt"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function fnDecrypt(cFeld) As String
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim cZeichen As String
    Dim lWert As Long
    Dim lWertMerker As Long
    Dim lZiel As Long
    Dim cZiel As String
    
    fnDecrypt = ""
    
    lWertMerker = 58
    For lcount = 1 To Len(cFeld)
        cZeichen = Mid(cFeld, lcount, 1)
        lWert = Asc(cZeichen)
        lZiel = lWert - lWertMerker
        If lZiel < 1 Then
            lZiel = (lWert + 255) - lWertMerker
        End If
        cZeichen = Chr$(lZiel)
        cZiel = cZiel & cZeichen
        lWertMerker = lZiel
    Next lcount
    
    'MsgBox cZiel
    
    fnDecrypt = cZiel
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "fndEcrypt"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ThisEineBedienerkarte(cstrichcode As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    ThisEineBedienerkarte = False
    
    sSQL = "select * from Bedname where bedcode = '" & Left(cstrichcode, 6) & Right(cstrichcode, 6) & "'"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!bedname) Then
            MsgBox "Möchten Sie Ihre Bedienerkarte verkaufen?", vbQuestion + vbYesNo, "Frage an: " & rsrs!bedname
            ThisEineBedienerkarte = True
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ThisEineBedienerkarte"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function fnPruefeEANWert(cEAN) As Integer
    On Error GoTo LOKAL_ERROR
    
    Dim lEANLen As Long
    Dim cZeichen As String
    Dim lPos As Long
    Dim lWert As Long
    Dim lSumme As Long
    
    fnPruefeEANWert = 0
    
    lEANLen = Len(cEAN)
    
    Select Case lEANLen
    
        Case Is = 6
        
        Case Is = 7
    
        Case Is = 8
            If Left(cEAN, 1) = "0" Then
            
            Else
                For lPos = 1 To 7
                    cZeichen = Mid(cEAN, lPos, 1)
                    If lPos / 2 = Int(lPos / 2) Then
                        lSumme = lSumme + Val(cZeichen)
                    Else
                        lSumme = lSumme + (Val(cZeichen) * 3)
                    End If
                Next lPos
                lWert = lSumme Mod 10
                If lWert > 0 Then
                    lWert = 10 - lWert
                End If
                If Trim$(Str$(lWert)) <> Mid(cEAN, 8, 1) Then
                    fnPruefeEANWert = 8
                End If
            End If
            
        Case Is = 10
        
            If Left(cEAN, 1) = "0" Then
            
            End If
            
        Case Is = 12
            For lPos = 1 To 11
                cZeichen = Mid(cEAN, lPos, 1)
                If lPos / 2 = Int(lPos / 2) Then
                    lSumme = lSumme + Val(cZeichen)
                Else
                    lSumme = lSumme + (Val(cZeichen) * 3)
                End If
            Next lPos
            lWert = lSumme Mod 10
            If lWert > 0 Then
                lWert = 10 - lWert
            End If
            If Trim$(Str$(lWert)) <> Mid(cEAN, 12, 1) Then
                fnPruefeEANWert = 12
            End If
        
        Case Is = 13
            For lPos = 1 To 12
                cZeichen = Mid(cEAN, lPos, 1)
                If lPos / 2 = Int(lPos / 2) Then
                    lSumme = lSumme + (Val(cZeichen) * 3)
                Else
                    lSumme = lSumme + Val(cZeichen)
                End If
            Next lPos
            lWert = lSumme Mod 10
            If lWert > 0 Then
                lWert = 10 - lWert
            End If
            If Trim$(Str$(lWert)) <> Mid(cEAN, 13, 1) Then
                fnPruefeEANWert = 13
            End If
        
        Case Else
            fnPruefeEANWert = 1
    End Select
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "fnPruefeEANWert"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Function fn_errechne_Prüfziffer(cEAN) As String
    On Error GoTo LOKAL_ERROR
    
    Dim lEANLen As Long
    Dim cZeichen As String
    Dim lPos As Long
    Dim lWert As Long
    Dim lSumme As Long
    
    fn_errechne_Prüfziffer = "x"
    
    lEANLen = Len(cEAN)
    
    Select Case lEANLen
        Case Is = 7
            If Left(cEAN, 1) = "0" Then
            
            Else
                For lPos = 1 To 7
                    cZeichen = Mid(cEAN, lPos, 1)
                    If lPos / 2 = Int(lPos / 2) Then
                        lSumme = lSumme + Val(cZeichen)
                    Else
                        lSumme = lSumme + (Val(cZeichen) * 3)
                    End If
                Next lPos
                lWert = lSumme Mod 10
                If lWert > 0 Then
                    lWert = 10 - lWert
                End If
                fn_errechne_Prüfziffer = lWert
            End If
            
        Case Is = 11
            For lPos = 1 To 11
                cZeichen = Mid(cEAN, lPos, 1)
                If lPos / 2 = Int(lPos / 2) Then
                    lSumme = lSumme + Val(cZeichen)
                Else
                    lSumme = lSumme + (Val(cZeichen) * 3)
                End If
            Next lPos
            lWert = lSumme Mod 10
            If lWert > 0 Then
                lWert = 10 - lWert
            End If
            fn_errechne_Prüfziffer = lWert
        Case Is = 12
            For lPos = 1 To 12
                cZeichen = Mid(cEAN, lPos, 1)
                If lPos / 2 = Int(lPos / 2) Then
                    lSumme = lSumme + (Val(cZeichen) * 3)
                Else
                    lSumme = lSumme + Val(cZeichen)
                End If
            Next lPos
            lWert = lSumme Mod 10
            If lWert > 0 Then
                lWert = 10 - lWert
            End If
            fn_errechne_Prüfziffer = lWert
    End Select
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "fn_errechne_Prüfziffer"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Function gbGrossLief(cLinr As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    gbGrossLief = False
    
    If NewTableSuchenDBKombi("GROLIEF", gdBase) Then
        If SpalteInTabellegefundenNEW("GROLIEF", "LINR", gdBase) = False Then
            loeschNEW "GROLIEF", gdBase
            CreateTable "GROLIEF", gdBase
        End If
    Else
        CreateTable "GROLIEF", gdBase
    End If
     
    cSQL = "Select * from GROLIEF where LINR = " & cLinr & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        gbGrossLief = True
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "gbGrossLief"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub PruefeRegistryEintragProgMOD01()
    On Error GoTo LOKAL_ERROR
    
    Dim cAppName As String      'KISS
    Dim cSection As String      'PROGRAMM
    Dim cKey As String          'WINKISS
    Dim cSetting As String      'PFAD
    
    cAppName = "KISS"
    cSection = "PROGRAMM"
    cKey = "WinKISS"
    
    cSetting = GetSetting(cAppName, cSection, cKey)
    
    If cSetting <> gcPfad Then
        cSetting = gcPfad
        SaveSetting cAppName, cSection, cKey, cSetting
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "PruefeRegistryEintragProgMOD01"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Sub PruefeRegistryEintragDataMOD01()
    On Error GoTo LOKAL_ERROR
    
    Dim cAppName As String      'KISS
    Dim cSection As String      'PROGRAMM
    Dim cKey As String          'WINKISS
    Dim cSetting As String      'PFAD
    
    cAppName = "KISS"
    cSection = "DATABASE"
    cKey = "WinKISS"
    
    cSetting = GetSetting(cAppName, cSection, cKey)
    
    If cSetting <> gcDBPfad Then
        cSetting = gcDBPfad
        SaveSetting cAppName, cSection, cKey, cSetting
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "PruefeRegistryEintragDataMOD01"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function ShortPath(ByVal Path As String) As String
On Error GoTo LOKAL_ERROR

    Dim Buffer As String * 255
    Dim rtn As Long
    
    If Len(Path) > 2 Then
        rtn = GetShortPathName(Path, Buffer, Len(Buffer))
        ShortPath = Left(Buffer, rtn)
    Else
        ShortPath = Path
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ShortPath"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function

Public Function fnBerechneMinuten(ctmp As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim cHH As String
    Dim cMM As String
    Dim cSS As String
    
    Dim lpos1 As Long
    Dim lPos2 As Long
    Dim lMin As Long
    
    fnBerechneMinuten = ""
    
    lpos1 = InStr(1, ctmp, ":")
    If lpos1 > 0 Then
        cHH = Left(ctmp, lpos1 - 1)
    End If
    
    lPos2 = InStr(lpos1 + 1, ctmp, ":")
    If lPos2 > 0 Then
        cMM = Mid(ctmp, lpos1 + 1, lPos2 - lpos1 - 1)
    Else
        cMM = Mid(ctmp, lpos1 + 1, Len(ctmp) - lpos1)
    End If
    
    If lPos2 > 0 Then
        cSS = Mid(ctmp, lPos2 + 1, Len(ctmp) - lPos2)
    End If
    
    lMin = (Val(cHH) * 60) + Val(cMM)
    
    fnBerechneMinuten = Trim$(Str$(lMin))
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "fnBerechneMinuten"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function fnPruefeUhrzeit(cUhrZeit) As Integer
    On Error GoTo LOKAL_ERROR
    
    'Uhrzeit wird im Format HH:MM erwartet
    'Rückgabewert ist 0 = okay oder 1 = Fehler
    
    Dim cZeichen As String
    Dim iCount As Integer
    Dim cValid As String
    Dim cStelle1 As String
    
    fnPruefeUhrzeit = 0
    
    If Len(cUhrZeit) <> 5 Then
        fnPruefeUhrzeit = 1
        Exit Function
    End If
    
    For iCount = 1 To 5
        cZeichen = Mid(cUhrZeit, iCount, 1)
        Select Case iCount
            Case Is = 1
                cValid = "012"
                If InStr(cValid, cZeichen) = 0 Then
                    fnPruefeUhrzeit = 1
                    Exit Function
                Else
                    cStelle1 = cZeichen
                End If
            Case Is = 2
                cValid = "0123456789"
                If InStr(cValid, cZeichen) = 0 Then
                    fnPruefeUhrzeit = 1
                    Exit Function
                Else
                    If cStelle1 = "2" Then
                        cValid = "0123"
                        If InStr(cValid, cZeichen) = 0 Then
                            fnPruefeUhrzeit = 1
                            Exit Function
                        End If
                    End If
                End If
            Case Is = 3
                cValid = ":"
                If InStr(cValid, cZeichen) = 0 Then
                    fnPruefeUhrzeit = 1
                    Exit Function
                End If
            Case Is = 4
                cValid = "012345"
                If InStr(cValid, cZeichen) = 0 Then
                    fnPruefeUhrzeit = 1
                    Exit Function
                End If
            Case Is = 5
                cValid = "0123456789"
                If InStr(cValid, cZeichen) = 0 Then
                    fnPruefeUhrzeit = 1
                    Exit Function
                End If
        End Select
                
    Next iCount
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "fnPruefeUhrzeit"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub LeseFirmenDaten()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim ctmp As String
    
    cSQL = "Select * from FIRMA"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!name) Then
            ctmp = rsrs!name
        Else
            ctmp = ""
        End If
        gFirma.FirmaName = ctmp
        If Not IsNull(rsrs!strasse) Then
            gFirma.strasse = rsrs!strasse
        Else
            gFirma.strasse = ""
        End If
        If Not IsNull(rsrs!Plz) Then
            gFirma.Plz = rsrs!Plz
        Else
            gFirma.Plz = ""
        End If
        If Not IsNull(rsrs!Ort) Then
            gFirma.Ort = rsrs!Ort
        Else
            gFirma.Ort = ""
        End If
        If Not IsNull(rsrs!Tel) Then
            gFirma.Tel = rsrs!Tel
        Else
            gFirma.Tel = ""
        End If
        If Not IsNull(rsrs!Fax) Then
            gFirma.Fax = rsrs!Fax
        Else
            gFirma.Fax = ""
        End If
        If Not IsNull(rsrs!BankName) Then
            gFirma.BankName = rsrs!BankName
        Else
            gFirma.BankName = ""
        End If
        If Not IsNull(rsrs!BLZ) Then
            gFirma.BLZ = rsrs!BLZ
        Else
            gFirma.BLZ = ""
        End If
        If Not IsNull(rsrs!Konto) Then
            gFirma.Konto = rsrs!Konto
        Else
            gFirma.Konto = ""
        End If
        If Not IsNull(rsrs!Steuernr) Then
            gFirma.Steuernr = rsrs!Steuernr
        Else
            gFirma.Steuernr = ""
        End If
        If Not IsNull(rsrs!ILN_1) Then
            gFirma.ILN_1 = rsrs!ILN_1
        Else
            gFirma.ILN_1 = ""
        End If
        If Not IsNull(rsrs!ILN_2) Then
            gFirma.ILN_2 = rsrs!ILN_2
        Else
            gFirma.ILN_2 = ""
        End If
        
        If Not IsNull(rsrs!BIC) Then
            gFirma.BIC = rsrs!BIC
        Else
            gFirma.BIC = ""
        End If
        If Not IsNull(rsrs!IBAN) Then
            gFirma.IBAN = rsrs!IBAN
        Else
            gFirma.IBAN = ""
        End If
        If Not IsNull(rsrs!Email) Then
            gFirma.FirmaMail = rsrs!Email
        Else
            gFirma.FirmaMail = ""
        End If
    Else
        gFirma.FirmaMail = ""
        gFirma.FirmaName = ""
        gFirma.strasse = ""
        gFirma.Plz = ""
        gFirma.Ort = ""
        gFirma.Tel = ""
        gFirma.Fax = ""
        gFirma.BankName = ""
        gFirma.BLZ = ""
        gFirma.Konto = ""
        gFirma.Steuernr = ""
        gFirma.ILN_1 = ""
        gFirma.ILN_2 = ""
        gFirma.BIC = ""
        gFirma.IBAN = ""
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "LeseFirmenDaten"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Public Sub LeseZugriffsRechte()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim dZugriff As Double
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim iFehler As Integer
    
    iFehler = 1
    
    For lcount = 0 To 32
        DlgZugriff(lcount).lcount = lcount
        DlgZugriff(lcount).dZugriff = 9
        DlgZugriff(lcount).dDlg = 0
    Next lcount
    
    iFehler = 2
    cSQL = "Select * from KISSLITE order by DIALOG "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    iFehler = 3
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            iFehler = 4
            If Not IsNull(rsrs!dialog) Then
                lcount = rsrs!dialog
            Else
                lcount = -1
            End If
            iFehler = 5

            If Not IsNull(rsrs!ZUGRIFF) Then
                dZugriff = rsrs!ZUGRIFF
            Else
                dZugriff = 9
            End If
            iFehler = 6

            Select Case lcount
                Case Is = 800

                    
                Case Is = 801

                    
                Case Is = 997
'                    If dZugriff = 1 Then
'                        gbRabatt = False
'                    Else
'                        gbRabatt = True
'                    End If
                    
                Case Is = 998

                    
                Case Is = 999

                    
                Case Else
                    DlgZugriff(lcount).dZugriff = dZugriff
                    
            End Select
            iFehler = 7
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "LeseZugriffsRechte"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Function ermittlezugriff(bytezugriffnr As Byte) As Byte
    On Error GoTo LOKAL_ERROR
    ermittlezugriff = 1
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    gsProteil = ""
    
    If bytezugriffnr <> 255 Then
        
        sSQL = "Select Proteil,Zugriff from BEDZUGRI where zugriffnr = " & bytezugriffnr
        
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            If Not IsNull(rsrs!ZUGRIFF) Then
                ermittlezugriff = rsrs!ZUGRIFF
            Else
                ermittlezugriff = 9
            End If
            
            If Not IsNull(rsrs!PROTEIL) Then
                gsProteil = rsrs!PROTEIL
            End If
        Else
            ermittlezugriff = 9
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermittlezugriff"
    Fehler.gsFehlertext = "Bei der Zugriffsermittlung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function LeseSpezpreis(lartnr As Long, bytePreistyp As Byte) As Single
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    LeseSpezpreis = 0

    sSQL = "Select * from Preise where ARTNR = " & lartnr
    sSQL = sSQL & " and Preistyp = " & bytePreistyp
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!Preiswert) Then
            LeseSpezpreis = rsrs!Preiswert
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "LeseSpezpreis"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function LeseStaffelpreis(lartnr As Long, lLinr As Long) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    LeseStaffelpreis = False

    sSQL = "Select * from STAFFELPR where ARTNR = " & lartnr & " and Linr = " & lLinr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        LeseStaffelpreis = True
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "LeseStaffelpreis"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function GibtEsMailText(lLinr As Long) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    GibtEsMailText = False

    sSQL = "Select * from LISRT_MAIL where LINR = " & lLinr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        GibtEsMailText = True
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "GibtEsMailText"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function LeseInterArt(cArtNr As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    LeseInterArt = False

    sSQL = "Select * from InterArt where ARTNR = " & cArtNr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        LeseInterArt = True
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "LeseInterArt"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function LeseGeschwisterArt(cArtNr As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    LeseGeschwisterArt = False
    
    If cArtNr = "" Then
        Exit Function
    End If
    
    If IsNumeric(cArtNr) = False Then
        Exit Function
    End If

    sSQL = "Select * from GESCHWART where mutterARTNR = " & cArtNr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        LeseGeschwisterArt = True
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "LeseGeschwisterArt"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function

Public Sub LeseStaffelpreisinList(lartnr As Long, Listx As ListBox, lLinr As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim rsrs    As Recordset
    Dim cFeld   As String
    Dim cSatz   As String
    
    Listx.Clear

    sSQL = "Select * from STAFFELPR where ARTNR = " & lartnr & " and LINR = " & lLinr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            cSatz = ""
            cFeld = ""
            If Not IsNull(rsrs!Menge) Then
                cFeld = rsrs!Menge
                cSatz = "ab " & cFeld
                If Not IsNull(rsrs!lekpr) Then
                    cFeld = rsrs!lekpr
                    cSatz = cSatz & " für " & Format(cFeld, "####0.00")
                    
                    If Not IsNull(rsrs!AENDER) Then
                        cFeld = rsrs!AENDER
                        cSatz = cSatz & " " & cFeld
                        
                        Listx.AddItem cSatz
                    End If
                End If
                
            End If
            
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "LeseStaffelpreisinList"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function LeseStaffelpreisArt(cArtNr As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    LeseStaffelpreisArt = False
    
    If cArtNr = "" Then
        Exit Function
    End If
    
    If IsNumeric(cArtNr) = False Then
        Exit Function
    End If

    sSQL = "Select * from STAFFELPRKVK where ARTNR = " & cArtNr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        LeseStaffelpreisArt = True
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "LeseStaffelpreisArt"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function LeseStaffelpreisArt_NEU(cArtNr As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    LeseStaffelpreisArt_NEU = False
    
    If gbMitStaffelPreis = False Then
        Exit Function
    End If
    
    If cArtNr = "" Then
        Exit Function
    End If
    
    If IsNumeric(cArtNr) = False Then
        Exit Function
    End If

    sSQL = "Select * from STAFFEL_KVK_ARTIKEL where ARTNR = " & cArtNr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        LeseStaffelpreisArt_NEU = True
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "LeseStaffelpreisArt_NEU"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermStaffNr(cArtNr As String) As Integer
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As DAO.Recordset
    
    ermStaffNr = 0
    
    If cArtNr = "" Then
        Exit Function
    End If
    
    If IsNumeric(cArtNr) = False Then
        Exit Function
    End If

    sSQL = "Select STAFFNR from STAFFEL_KVK_ARTIKEL where ARTNR = " & cArtNr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
    
        If Not IsNull(rsrs!STAFFNR) Then
            ermStaffNr = rsrs!STAFFNR
                        
        End If
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermStaffNr"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function


Public Sub LeseVedesFTP()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim rsrs    As Recordset
    
    If NewTableSuchenDBKombi("VEDESFTP", gdBase) = False Then
        CreateTableT2 "VEDESFTP", gdBase
    End If
    

    sSQL = "Select * from VEDESFTP "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!Host) Then
            gsVEDES_HOST = rsrs!Host
        End If
        
        If Not IsNull(rsrs!User) Then
            gsVEDES_USER = rsrs!User
        End If
        
        If Not IsNull(rsrs!PW) Then
            gsVEDES_PW = rsrs!PW
        End If
            
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "LeseVedesFTP"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub LeseVedesFTP_DSL()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim rsrs    As Recordset
    
    If NewTableSuchenDBKombi("VEDESFTP_DSL", gdBase) = False Then
        CreateTableT2 "VEDESFTP_DSL", gdBase
    End If
    

    sSQL = "Select * from VEDESFTP_DSL "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!Host) Then
            gsVEDES_HOST_DSL = rsrs!Host
        End If
        
        If Not IsNull(rsrs!User) Then
            gsVEDES_USER_DSL = rsrs!User
        End If
        
        If Not IsNull(rsrs!PW) Then
            gsVEDES_PW_DSL = rsrs!PW
        End If
            
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "LeseVedesFTP_DSL"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub LeseTSE_EINSTELLUNG()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim rsrs    As DAO.Recordset
    
    If NewTableSuchenDBKombi("TSE_ONLEINSTELLUNG", gdApp) = False Then
        CreateTableT3 "TSE_ONLEINSTELLUNG", gdApp
    End If
    
    gbTSE_SCHREIBEN = False
    gsTSE_APIKEY = ""
    gsTSE_APISECRET = ""
    gsTSE_TSEID = ""
    gsTSE_CLIENTID = ""
    
    sSQL = "Select * from TSE_ONLEINSTELLUNG "
    Set rsrs = gdApp.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!TSE_SCHREIBEN) Then
            gbTSE_SCHREIBEN = rsrs!TSE_SCHREIBEN
        End If
        
        If Not IsNull(rsrs!APIKEY) Then
            gsTSE_APIKEY = rsrs!APIKEY
        End If
        
        If Not IsNull(rsrs!APISECRET) Then
            gsTSE_APISECRET = rsrs!APISECRET
        End If
        
        If Not IsNull(rsrs!TSEID) Then
            gsTSE_TSEID = rsrs!TSEID
        End If
        
        If Not IsNull(rsrs!clientID) Then
            gsTSE_CLIENTID = rsrs!clientID
        End If
        
        
    
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "LeseTSE_EINSTELLUNG"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub LeseMehrEAN(lartnr As Long, Listx As ListBox)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim rsrs    As Recordset
    Dim cFeld   As String
    Dim cSatz   As String
    
    Listx.Clear

    sSQL = "Select * from ARTEAN_K where ARTNR = " & lartnr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            cSatz = ""
            cFeld = ""
            If Not IsNull(rsrs!EAN) Then
                cFeld = rsrs!EAN
                cSatz = cFeld
                Listx.AddItem cSatz
            End If
            
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "LeseMehrEAN"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function StaffelKVK_vorhanden(lartnr As Long) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim rsrs    As Recordset
    Dim cFeld   As String
    Dim cSatz   As String
    
    StaffelKVK_vorhanden = False

    sSQL = "Select * from Staffel_KVK_Artikel where ARTNR = " & lartnr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        StaffelKVK_vorhanden = True
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "StaffelKVK_vorhanden"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function MehrEAN_vorhanden(lartnr As Long) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim rsrs    As Recordset
    Dim cFeld   As String
    Dim cSatz   As String
    
    MehrEAN_vorhanden = False

    sSQL = "Select * from ARTEAN_K where ARTNR = " & lartnr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        MehrEAN_vorhanden = True
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "MehrEAN_vorhanden"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function fnMoveComma2Point$(ctmp As String)
    On Error GoTo LOKAL_ERROR
    
    ctmp = Trim$(ctmp)
    
    If InStr(ctmp, ".") Then
        ctmp = fnEntfernePunkt$(ctmp)
    End If
    
    If InStr(ctmp, ",") > 0 Then
        Mid(ctmp, InStr(ctmp, ","), 1) = "."
    End If
    
    fnMoveComma2Point$ = ctmp
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "fnMoveComma2Point$"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function


Private Function fnEntfernePunkt$(ctmp As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cZiel As String
    Dim lcount As Long
    
    For lcount = 1 To Len(ctmp)
        cZeichen = Mid(ctmp, lcount, 1)
        If cZeichen <> "." Then
            cZiel = cZiel & cZeichen
        End If
    Next lcount
    
    fnEntfernePunkt$ = cZiel
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "fnEntfernePunkt$"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub KonvertAnsiAscii(cText As String)
    On Error GoTo LOKAL_ERROR
    
    Dim lPos As Long
    
    If gbSPY Then
        frmWKL20!Winsock22.senddata cText & vbCrLf
    End If

    Do While InStr(cText, Chr$(196)) > 0
        lPos = InStr(cText, Chr$(196))
        Mid(cText, lPos, 1) = Chr$(142)
    Loop
    
    Do While InStr(cText, Chr$(214)) > 0
        lPos = InStr(cText, Chr$(214))
        Mid(cText, lPos, 1) = Chr$(153)
    Loop
    
    Do While InStr(cText, Chr$(220)) > 0
        lPos = InStr(cText, Chr$(220))
        Mid(cText, lPos, 1) = Chr$(154)
    Loop
    
    Do While InStr(cText, Chr$(228)) > 0
        lPos = InStr(cText, Chr$(228))
        Mid(cText, lPos, 1) = Chr$(132)
    Loop
    
    Do While InStr(cText, Chr$(246)) > 0
        lPos = InStr(cText, Chr$(246))
        Mid(cText, lPos, 1) = Chr$(148)
    Loop
    
    Do While InStr(cText, Chr$(252)) > 0
        lPos = InStr(cText, Chr$(252))
        Mid(cText, lPos, 1) = Chr$(129)
    Loop
    
    Do While InStr(cText, Chr$(223)) > 0
        lPos = InStr(cText, Chr$(223))
        Mid(cText, lPos, 1) = Chr$(225)
    Loop
    
    Do While InStr(cText, Chr$(233)) > 0
        lPos = InStr(cText, Chr$(233))
        Mid(cText, lPos, 1) = Chr$(130)
    Loop
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "KonvertAnsiAscii"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Public Sub KonvertAsciiAnsi(cText As String)
    On Error GoTo LOKAL_ERROR
    
    Dim lPos As Long
    
    Do While InStr(cText, Chr$(142)) > 0
        lPos = InStr(cText, Chr$(142))
        Mid(cText, lPos, 1) = Chr$(196)
    Loop
    
    Do While InStr(cText, Chr$(153)) > 0
        lPos = InStr(cText, Chr$(153))
        Mid(cText, lPos, 1) = Chr$(214)
    Loop
    
    Do While InStr(cText, Chr$(154)) > 0
        lPos = InStr(cText, Chr$(154))
        Mid(cText, lPos, 1) = Chr$(220)
    Loop
    
    Do While InStr(cText, Chr$(132)) > 0
        lPos = InStr(cText, Chr$(132))
        Mid(cText, lPos, 1) = Chr$(228)
    Loop
    
    Do While InStr(cText, Chr$(148)) > 0
        lPos = InStr(cText, Chr$(148))
        Mid(cText, lPos, 1) = Chr$(246)
    Loop
    
    Do While InStr(cText, Chr$(129)) > 0
        lPos = InStr(cText, Chr$(129))
        Mid(cText, lPos, 1) = Chr$(252)
    Loop
    
    Do While InStr(cText, Chr$(225)) > 0
        lPos = InStr(cText, Chr$(225))
        Mid(cText, lPos, 1) = Chr$(223)
    Loop
    
    Do While InStr(cText, Chr$(130)) > 0
        lPos = InStr(cText, Chr$(130))
        Mid(cText, lPos, 1) = Chr$(233)
    Loop
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "KonvertAsciiAnsi"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function ermFarbkz(cFarbnr As String) As String
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rs As Recordset
    
    ermFarbkz = ""
    
    sSQL = "Select Farbtext From FARBMERK "
    sSQL = sSQL & "  where FARBNR = " & cFarbnr

    Set rs = gdBase.OpenRecordset(sSQL)
    
    If Not rs.EOF Then
    rs.MoveFirst
        If Not IsNull(rs!farbtext) Then
            ermFarbkz = rs!farbtext
        End If
    End If
    rs.Close
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermFarbkz"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function ermFarbe(cKdnr As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    cSQL = "Select * from KUNDEN where KUNDNR = " & cKdnr & " "
    FnOpenrecordset rsrs, cSQL, 1, gdBase
    
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!AWM) Then
            ermFarbe = rsrs!AWM
        Else
            ermFarbe = "0"
        End If
    Else
        ermFarbe = "0"
    End If
    
    If IsNumeric(ermFarbe) = False Then
        ermFarbe = "0"
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermfarbe"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function Ermittleagntext(sZiff As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsArt As Recordset
  
    Ermittleagntext = ""
    sSQL = "Select * from agndbf where agn = " & Val(sZiff)
    
    Set rsArt = gdBase.OpenRecordset(sSQL)
    If Not rsArt.EOF Then
        Ermittleagntext = rsArt!AGTEXT
    Else
        Ermittleagntext = "Artikelgruppe ist nicht definiert"
    End If
    
    rsArt.Close: Set rsArt = Nothing
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "Ermittleagntext"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ErmittleGruppenbez(sZiff As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsArt As Recordset
  
    ErmittleGruppenbez = ""
    sSQL = "Select * from GRUPPE where GRUPPENNR = " & Val(sZiff)
    
    Set rsArt = gdBase.OpenRecordset(sSQL)
    If Not rsArt.EOF Then
        ErmittleGruppenbez = rsArt!Gruppenbez
    Else
        ErmittleGruppenbez = "Gruppe ist nicht definiert"
    End If
    
    rsArt.Close: Set rsArt = Nothing
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ErmittleGruppenbez"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function Ermittlepgntext(sZiff As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsArt As Recordset
  
    Ermittlepgntext = ""
    sSQL = "Select * from pgndbf where pgn = " & Val(sZiff)
    
    Set rsArt = gdBase.OpenRecordset(sSQL)
    If Not rsArt.EOF Then
        Ermittlepgntext = rsArt!PGNBEZEICH
    Else
        Ermittlepgntext = "Produktgruppe ist nicht definiert"
    End If
    
    rsArt.Close: Set rsArt = Nothing
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "Ermittlepgntext"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub setzeFarbeinWK(lartikel As Long, sawm As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim rsrs    As Recordset
    
    MerkeFarbeVorher lartikel
    
    sSQL = "Select awm from artikel where artnr = " & lartikel
    Set rsrs = gdBase.OpenRecordset(sSQL)
   
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        rsrs.Edit
        rsrs!AWM = sawm
        rsrs.Update
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "setzeFarbeinWK"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Public Sub MerkeFarbeVorher(lartikel As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    
    sSQL = "Delete from MERKFARB where artnr = " & lartikel
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into  MERKFARB select artnr , awm from artikel "
    sSQL = sSQL & " where artnr = " & lartikel
    sSQL = sSQL & " and trim(awm) <> '0' "
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "MerkeFarbeVorher"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Public Function ermMerkFarbe(sartikel As String, sawmkrti As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim rsrs    As Recordset
    Dim rsArt   As Recordset
    
    ermMerkFarbe = "0"
    
    'achtung nur wenn farbe jetzt auf 95
    
    sSQL = "select * from Artikel where artnr = " & sartikel
    sSQL = sSQL & " and awm = '" & sawmkrti & "'"
    Set rsArt = gdBase.OpenRecordset(sSQL)
    If Not rsArt.EOF Then
    
        sSQL = "select * from MERKFARB where artnr = " & sartikel
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!AWM) Then
                ermMerkFarbe = rsrs!AWM
            End If
        
        End If
        rsrs.Close: Set rsrs = Nothing
        
    End If
    rsArt.Close: Set rsArt = Nothing
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermMerkFarbe"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermStichtag() As Date
    On Error GoTo LOKAL_ERROR
    
    ermStichtag = 0
    
    Dim cSQL As String
    Dim rsINB As Recordset
    
    cSQL = "Select Datum from Stichtag  "
    Set rsINB = gdBase.OpenRecordset(cSQL)
    If Not rsINB.EOF Then
        If Not IsNull(rsINB!Datum) Then
            ermStichtag = rsINB!Datum
        End If
    End If
    rsINB.Close
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermStichtag"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function ermGutscheinAusgabeTag(sGutschnr As String) As Date
    On Error GoTo LOKAL_ERROR
    
    ermGutscheinAusgabeTag = 0
    
    Dim cSQL As String
    Dim rsINB As Recordset
    
    If gbKL_LIVEGUTSCHEIN Then
    
    
            If fTestLogin_SQLDABA_Error = 0 Then 'ist alles OK? Datenbank erreichbar?
            'alles okay
        Else
            schreibeProtokollVPNTXT "Unterbrechung"
            
            Dim sTemp As String
            sTemp = "Bitte starten Sie diesen Rechner neu" & vbCrLf
            sTemp = sTemp & "oder schließen Sie das Schloss und starten Sie WinKiss neu."
        
            MsgBox sTemp, vbCritical + vbOKOnly, "Gutschein-Datenbank nicht erreichbar"
            Exit Function
        End If
        
        
        Dim stConnect As String
    
        If gsKL_DSN <> "" Then
            stConnect = "ODBC;DSN=" & gsKL_DSN & ";UID=" & gsKL_BENUTZER & ";PWD=" & gsKL_PASSWORT & ""
        Else
            stConnect = "ODBC;DRIVER=SQL Server;SERVER=" & gsKL_ADRESSE & ";DATABASE=" & gsKL_DATENBANKNAME & ";UID=" & gsKL_BENUTZER & ";PWD=" & gsKL_PASSWORT & ""
        End If
        
        Dim dbEAN As DAO.Database
        Set dbEAN = OpenDatabase(gsKL_DATENBANKNAME, dbDriverNoPrompt, False, stConnect)
        
        
        cSQL = "Select * from GUTSCHEINE where GUTSCHNR = '" & sGutschnr & "'"
        Set rsINB = dbEAN.OpenRecordset(cSQL)
        If Not rsINB.EOF Then
            If Not IsNull(rsINB!AUSG_DATUM) Then
                
                ermGutscheinAusgabeTag = rsINB!AUSG_DATUM
            End If
        End If
        rsINB.Close: Set rsINB = Nothing
        
        
        
        dbEAN.Close
    
    
    
    
    Else
    
        
        
        cSQL = "Select DAT_AUSG from GUTSCH where gutschnr = " & sGutschnr
        Set rsINB = gdBase.OpenRecordset(cSQL)
        If Not rsINB.EOF Then
            If Not IsNull(rsINB!DAT_AUSG) Then
                ermGutscheinAusgabeTag = rsINB!DAT_AUSG
            End If
        End If
        rsINB.Close
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermGutscheinAusgabeTag"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Sub DELMerkFarbe(sartikel As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    
    sSQL = "Delete from MERKFARB where artnr = " & sartikel
    gdBase.Execute sSQL, dbFailOnError
         
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "DELMerkFarbe"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub lese_Termin_Optionen()
On Error GoTo LOKAL_ERROR

    Dim sSQL    As String
    Dim rsrs    As Recordset
    
    gbTerm_Name = False
    gbTerm_InfoDauerh = False
    gbTerm_BedKass = False
    
    If NewTableSuchenDBKombi("TermOptionen", gdBase) Then
    
        sSQL = "select * from TermOptionen "
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!Term_Name) Then
                gbTerm_Name = rsrs!Term_Name
            End If
            
            If Not IsNull(rsrs!Term_InfoDauerh) Then
                gbTerm_InfoDauerh = rsrs!Term_InfoDauerh
            End If
            
            If Not IsNull(rsrs!Term_BedKass) Then
                gbTerm_BedKass = rsrs!Term_BedKass
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "lese_Termin_Optionen"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Public Function ErmlzVKproFil(cART As String, iFil As Integer) As Date
    On Error GoTo LOKAL_ERROR
    
    ErmlzVKproFil = 0
    
    Dim cSQL As String
    Dim rsINB As Recordset
    
    cSQL = "Select max(adate) as maxdate from Kassjour where ARTNR = " & cART & " "
    cSQL = cSQL & " and Filiale = " & iFil
    Set rsINB = gdBase.OpenRecordset(cSQL)
    If Not rsINB.EOF Then
        If Not IsNull(rsINB!MaxDate) Then
            ErmlzVKproFil = rsINB!MaxDate
        End If
    End If
    rsINB.Close
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ErmlzVKproFil"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function KomprimieredieAPP() As Boolean
On Error GoTo LOKAL_ERROR

    Dim dFilesize As Double
    Dim i As Integer
    dFilesize = FileLen(App.Path & "\kissapp.mdb")     'in BYTE
    dFilesize = dFilesize / 1024                    'in KBYTE
    dFilesize = dFilesize / 1024

    If dFilesize > 400 Then
'        lblA.Visible = True
'        anzeige "normal", "Bitte warten, Datenbank wird aufgeräumt... ", lblA
        loeschNEW "KASSJOUR", gdApp
        loeschNEW "KASS", gdApp
        loeschNEW "ZBESTAND", gdApp
        loeschNEW "ARTIKEL", gdApp
        loeschNEW "FU4", gdApp
        loeschNEW "LAGERPLATZ", gdApp
        loeschNEW "LINBEZ", gdApp
        loeschNEW "LISRT", gdApp
        
        For i = 0 To 15
            loeschNEW "KASS" & i, gdApp
        Next i
        
        If BistDualleineinderDatenbankApp Then
            dbApp_Compri "Kissapp.MDB"
            
        End If
'        lblA.Visible = False
    End If

Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "KomprimieredieAPP"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Sub AlleZugriffeLöschen()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    sSQL = "Delete from ZugriffDat where rechner = '" & srechnertab & "'"
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "AlleZugriffeLöschen"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Public Sub insert_BestVor(sArtnr As String, Optional sLinr As String = "")
    On Error GoTo LOKAL_ERROR
    
    If sArtnr = "" Then
        Exit Sub
    End If
    
    If IsNumeric(sArtnr) = False Then
        Exit Sub
    End If
    
    
    
    
    
    
    
    
    
    Dim cSQL As String
    
    If NewTableSuchenDBKombi("X" & sLinr & "_BV", gdBase) = False Then
    
        cSQL = "Create Table X" & sLinr & "_BV"
        cSQL = cSQL & "( ARTNR float"
        cSQL = cSQL & ", BEZEICH varchar(35)"
        cSQL = cSQL & ", AGN float"
        cSQL = cSQL & ", LINR float"
        cSQL = cSQL & ", PGN float"
        cSQL = cSQL & ", LIBESNR varchar(13)"
        cSQL = cSQL & ", EAN varchar(13)"
        cSQL = cSQL & ", RKZ varchar(1)"
        cSQL = cSQL & ", LPZ float"
        cSQL = cSQL & ", AWM varchar(2)"
        cSQL = cSQL & ", AWM2 varchar(2)"
        cSQL = cSQL & ", MERK varchar(4)"
        cSQL = cSQL & ", EKPR float"
        cSQL = cSQL & ", LEKPR float"
        cSQL = cSQL & ", VKPR float"
        cSQL = cSQL & ", KVKPR1 float"
        cSQL = cSQL & ", MOPREIS float"
        cSQL = cSQL & ", MINMEN float"
        cSQL = cSQL & ", MINBEST int"
        cSQL = cSQL & ", BESTVOR float"
        cSQL = cSQL & ", SOKO int"
        cSQL = cSQL & ", FAKTOR float"
        cSQL = cSQL & ", LPZ_VON int"
        cSQL = cSQL & ", LPZ_BIS int"
        cSQL = cSQL & ", EINDECK int"
        cSQL = cSQL & ", BEVORRAT float"
        cSQL = cSQL & ", INBEST int"
        cSQL = cSQL & ", BESTAND float"
        cSQL = cSQL & ", VKAMo1 float"
        cSQL = cSQL & ", VKVMo1 float"
        cSQL = cSQL & ", VKLJ1 float"
        cSQL = cSQL & ", VKVJ1 float"
        cSQL = cSQL & ", MITTEILUNG varchar(250)"
        cSQL = cSQL & ", ANZEIGE smallint"
        cSQL = cSQL & ", NOTIZEN varchar(25)"
        cSQL = cSQL & ", LJ1 smallint"
        cSQL = cSQL & ", LJ2 smallint"
        cSQL = cSQL & ", LJ3 smallint"
        cSQL = cSQL & ", LJ4 smallint"
        cSQL = cSQL & ", LJ5 smallint"
        cSQL = cSQL & ", LJ6 smallint"
        cSQL = cSQL & ", LJ7 smallint"
        cSQL = cSQL & ", LJ8 smallint"
        cSQL = cSQL & ", LJ9 smallint"
        cSQL = cSQL & ", LJ10 smallint"
        cSQL = cSQL & ", LJ11 smallint"
        cSQL = cSQL & ", LJ12 smallint"
        cSQL = cSQL & ", VJ1 smallint"
        cSQL = cSQL & ", VJ2 smallint"
        cSQL = cSQL & ", VJ3 smallint"
        cSQL = cSQL & ", VJ4 smallint"
        cSQL = cSQL & ", VJ5 smallint"
        cSQL = cSQL & ", VJ6 smallint"
        cSQL = cSQL & ", VJ7 smallint"
        cSQL = cSQL & ", VJ8 smallint"
        cSQL = cSQL & ", VJ9 smallint"
        cSQL = cSQL & ", VJ10 smallint"
        cSQL = cSQL & ", VJ11 smallint"
        cSQL = cSQL & ", VJ12 smallint"
        cSQL = cSQL & ", SORTI smallint"
        cSQL = cSQL & ", GROESSE varchar(10)"
        cSQL = cSQL & ", PIN varchar(1)"
        cSQL = cSQL & ", SHOP varchar(1)"
        cSQL = cSQL & ") "
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    Dim rsrs As DAO.Recordset

    Dim bUpdaten As Boolean
    bUpdaten = False


    cSQL = "Select * from X" & sLinr & "_BV where ARTNR = " & sArtnr
            
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        bUpdaten = True
    Else
    
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If bUpdaten = True Then
    
        cSQL = "Update X" & sLinr & "_BV set BESTVOR = BESTVOR + 1 "
        cSQL = cSQL & " where artnr = " & sArtnr
        gdBase.Execute cSQL, dbFailOnError
    
    Else
    

    
        cSQL = "Insert into X" & sLinr & "_BV select "
        cSQL = cSQL & " a.ARTNR "
        cSQL = cSQL & ", a.BEZEICH "
        cSQL = cSQL & ", a.AGN "
        cSQL = cSQL & ", " & sLinr & " as LINR "
        cSQL = cSQL & ", a.PGN "
        cSQL = cSQL & ", b.LIBESNR "
        cSQL = cSQL & ", a.EAN "
        cSQL = cSQL & ", b.RKZ "
        cSQL = cSQL & ", a.LPZ "
        cSQL = cSQL & ", a.AWM "
        cSQL = cSQL & ", '' as AWM2 "
        cSQL = cSQL & ", '' as MERK "
        cSQL = cSQL & ", a.EKPR "
        cSQL = cSQL & ", b.LEKPR "
        cSQL = cSQL & ", a.VKPR "
        cSQL = cSQL & ", a.KVKPR1 "
        cSQL = cSQL & ", 0 as MOPREIS "
        cSQL = cSQL & ", b.MINMEN "
        cSQL = cSQL & ", a.MINBEST "
        cSQL = cSQL & ", 1 as BESTVOR "
        cSQL = cSQL & ", 0 as SOKO"
        cSQL = cSQL & ", 0 as FAKTOR "
        cSQL = cSQL & ", 1 as LPZ_VON "
        cSQL = cSQL & ", 999 as LPZ_BIS "
        cSQL = cSQL & ", 1 as EINDECK "
        cSQL = cSQL & ", 1 as BEVORRAT "
        cSQL = cSQL & ", 0 as INBEST "
        cSQL = cSQL & ", a.BESTAND "
        cSQL = cSQL & ", 0 as VKAMo1 "
        cSQL = cSQL & ", 0 as VKVMo1 "
        cSQL = cSQL & ", 0 as VKLJ1 "
        cSQL = cSQL & ", 0 as VKVJ1 "
        cSQL = cSQL & ", '' as  MITTEILUNG "
        cSQL = cSQL & ", 0 as ANZEIGE "
        cSQL = cSQL & ", '' as NOTIZEN "
        cSQL = cSQL & ", 0 as LJ1 "
        cSQL = cSQL & ", 0 as LJ2 "
        cSQL = cSQL & ", 0 as LJ3 "
        cSQL = cSQL & ", 0 as LJ4 "
        cSQL = cSQL & ", 0 as LJ5 "
        cSQL = cSQL & ", 0 as LJ6 "
        cSQL = cSQL & ", 0 as LJ7 "
        cSQL = cSQL & ", 0 as LJ8 "
        cSQL = cSQL & ", 0 as LJ9 "
        cSQL = cSQL & ", 0 as LJ10 "
        cSQL = cSQL & ", 0 as LJ11 "
        cSQL = cSQL & ", 0 as LJ12 "
        cSQL = cSQL & ", 0 as VJ1 "
        cSQL = cSQL & ", 0 as VJ2 "
        cSQL = cSQL & ", 0 as VJ3 "
        cSQL = cSQL & ", 0 as VJ4 "
        cSQL = cSQL & ", 0 as VJ5 "
        cSQL = cSQL & ", 0 as VJ6 "
        cSQL = cSQL & ", 0 as VJ7 "
        cSQL = cSQL & ", 0 as VJ8 "
        cSQL = cSQL & ", 0 as VJ9 "
        cSQL = cSQL & ", 0 as VJ10 "
        cSQL = cSQL & ", 0 as VJ11 "
        cSQL = cSQL & ", 0 as VJ12 "
        cSQL = cSQL & ", 0 as SORTI "
        cSQL = cSQL & ", '' as GROESSE "
        cSQL = cSQL & ", '' as PIN "
        cSQL = cSQL & ", '' as SHOP "
        cSQL = cSQL & " from artikel a inner join artlief b on a.artnr = b.artnr where b.linr = " & sLinr & " "
        cSQL = cSQL & " and a.artnr = " & sArtnr
        gdBase.Execute cSQL, dbFailOnError
        
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "insert_BestVor"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Public Function ermINBV(sArtnr As String, Optional sLinr As String = "") As Long
    On Error GoTo LOKAL_ERROR
    
    ermINBV = 0
    
    If sArtnr = "" Then
        Exit Function
    End If
    
    If IsNumeric(sArtnr) = False Then
        Exit Function
    End If

    Dim cSQL As String
    Dim rsrs As DAO.Recordset
    
    'alle Tabellen mit X abklappern
    
    If sLinr <> "" Then
        If NewTableSuchenDBKombi("X" & sLinr & "_BV", gdBase) = True Then
        
            cSQL = "Select BESTVOR from X" & sLinr & "_BV where ARTNR = " & sArtnr
            
            Set rsrs = gdBase.OpenRecordset(cSQL)
            If Not rsrs.EOF Then
                rsrs.MoveFirst
                If Not IsNull(rsrs!BESTVOR) Then
                    ermINBV = rsrs!BESTVOR
                End If
            End If
            rsrs.Close: Set rsrs = Nothing
        
        End If
    
'        If NewTableSuchenDBKombi("X" & sLinr & "", gdBase) = False Then
'
'        End If
       
    End If
    
    
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermINBV"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Function checkFFE() As Boolean
    On Error GoTo LOKAL_ERROR
    
    checkFFE = False
    
    If NewTableSuchenDBKombi("FFE", gdBase) Then
        If Not SpalteInTabellegefundenNEW("FFE", "AGNHOFF", gdBase) Then
            SpalteAnfuegenNEW "FFE", "AGNHOFF", "LONG", gdBase
            SpalteAnfuegenNEW "FFE", "AGNVEDE", "LONG", gdBase
        End If
        
        If Not SpalteInTabellegefundenNEW("FFE", "Hoffdoppel", gdBase) Then
            SpalteAnfuegenNEW "FFE", "Hoffdoppel", "BIT", gdBase
            SpalteAnfuegenNEW "FFE", "Vededoppel", "BIT", gdBase
        End If
        
        If Not SpalteInTabellegefundenNEW("FFE", "Hoffhof", gdBase) Then
            SpalteAnfuegenNEW "FFE", "Hoffhof", "BIT", gdBase
        End If
        
        If Not SpalteInTabellegefundenNEW("FFE", "RAB", gdBase) Then
            SpalteAnfuegenNEW "FFE", "RAB", "BIT", gdBase
        End If
        
        If Not SpalteInTabellegefundenNEW("FFE", "AGN", gdBase) Then
            SpalteAnfuegenNEW "FFE", "AGN", "LONG", gdBase
        End If
        
        If Not SpalteInTabellegefundenNEW("FFE", "rewe", gdBase) Then
            SpalteAnfuegenNEW "FFE", "rewe", "BIT", gdBase
        End If
        
        If Not SpalteInTabellegefundenNEW("FFE", "LUENING", gdBase) Then
            SpalteAnfuegenNEW "FFE", "LUENING", "BIT", gdBase
        End If
        
        If Not SpalteInTabellegefundenNEW("FFE", "agnTUBZigaretten", gdBase) Then
            SpalteAnfuegenNEW "FFE", "agnTUBZigaretten", "LONG", gdBase
        End If
        
        If Not SpalteInTabellegefundenNEW("FFE", "agnTUBZigarren", gdBase) Then
            SpalteAnfuegenNEW "FFE", "agnTUBZigarren", "LONG", gdBase
        End If
        
        If Not SpalteInTabellegefundenNEW("FFE", "agnTUBTabak", gdBase) Then
            SpalteAnfuegenNEW "FFE", "agnTUBTabak", "LONG", gdBase
        End If
        
        If Not SpalteInTabellegefundenNEW("FFE", "agnTUBFeinschnitt", gdBase) Then
            SpalteAnfuegenNEW "FFE", "agnTUBFeinschnitt", "LONG", gdBase
        End If
        
        If Not SpalteInTabellegefundenNEW("FFE", "agnTUBPfeifentabak", gdBase) Then
            SpalteAnfuegenNEW "FFE", "agnTUBPfeifentabak", "LONG", gdBase
        End If
        
        If Not SpalteInTabellegefundenNEW("FFE", "agnTUBRBA", gdBase) Then
            SpalteAnfuegenNEW "FFE", "agnTUBRBA", "LONG", gdBase
        End If
        
        If Not SpalteInTabellegefundenNEW("FFE", "AGNZEITUNGe", gdBase) Then
            SpalteAnfuegenNEW "FFE", "AGNZEITUNGe", "LONG", gdBase
        End If
        
        If Not SpalteInTabellegefundenNEW("FFE", "AGNZEITUNGv", gdBase) Then
            SpalteAnfuegenNEW "FFE", "AGNZEITUNGv", "LONG", gdBase
        End If
        
        If Not SpalteInTabellegefundenNEW("FFE", "EANNULLEN", gdBase) Then
            SpalteAnfuegenNEW "FFE", "EANNULLEN", "BIT", gdBase
        End If
        
        If Not SpalteInTabellegefundenNEW("FFE", "VERGLEICH", gdBase) Then
            SpalteAnfuegenNEW "FFE", "VERGLEICH", "integer", gdBase
        End If
        
        If Not SpalteInTabellegefundenNEW("FFE", "agnGerry", gdBase) Then
            SpalteAnfuegenNEW "FFE", "agnGerry", "LONG", gdBase
        End If
        
        If Not SpalteInTabellegefundenNEW("FFE", "agnLue", gdBase) Then
            SpalteAnfuegenNEW "FFE", "agnLue", "LONG", gdBase
        End If
        
        If Not SpalteInTabellegefundenNEW("FFE", "KEINDEL", gdBase) Then
            SpalteAnfuegenNEW "FFE", "KEINDEL", "BIT", gdBase
        End If
        
        checkFFE = True
    Else
        checkFFE = False
    End If
    
     Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "checkFFE"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
