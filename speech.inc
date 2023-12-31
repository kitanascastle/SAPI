#IF NOT %DEF(%WIN32API_INC)
  'PB WinApi verwenden
  #INCLUDE ONCE "WIN32API.INC"
  #INCLUDE ONCE "SAPI.INC"
  $CLSID_SPEECHLIB = $CLSID_SpeechLib_SpVoice
  $CLSID_TOTEXTLIB = $CLSID_SpSharedRecognizer
#ELSE
  'Jose Roca WinApi verwenden
  #INCLUDE ONCE "SAPI.INC"
  $CLSID_SPEECHLIB = $CLSID_SpVoice
  $CLSID_TOTEXTLIB = $CLSID_SpSharedRecognizer
#ENDIF

%SPF_DEFAULT = 0
%SPF_ASYNC = 1
%SPF_PURGEBEFORESPEAK = 2

%GRAMMAR_ID = 0
$GRAMMAR_RULE = "RULE1"
%MAX_TEXTRECOGNITION_EVENTS = 10

%SAPIERR_FAILED_TO_CREATE_SPVOICE = -2
%SAPIERR_FAILED_TO_CREATE_RECOGNIZER = -3
%SAPIERR_FAILED_TO_CREATE_CONTEXT = -4
%SAPIERR_FAILED_TO_CREATE_GRAMMAR = -5
%SAPIERR_FAILED_TO_RESET_GRAMMAR = -6
%SAPIERR_FAILED_TO_GET_RULE = -7
%SAPIERR_FAILED_TO_SET_NOTIFY = -8
%SAPIERR_FAILED_TO_SET_INTEREST = -9
%SAPIERR_FAILED_TO_ACTIVATE_RULE = -10
%SAPIERR_FAILED_TO_RESUME_CONTEXT = -11

GLOBAL pISpVoice AS ISpVoice
GLOBAL pITextRecognizer AS ISpRecognizer
GLOBAL pIRecoContext AS ISpRecoContext
GLOBAL pIRecoGrammar AS ISpRecoGrammar
GLOBAL SAPIINSTALLED&
GLOBAL SAPINOWAIT&
GLOBAL STT_NOTIFICATION_HANDLE&



'Fehlertext zu Fehlercode ermitteln
FUNCTION SAPIGETERRORMSG$(ERRCODE&)
  LOCAL A$

  SELECT CASE ERRCODE&
  CASE %SAPIERR_FAILED_TO_CREATE_SPVOICE: A$ = "Failed to create SpVoice"
  CASE %SAPIERR_FAILED_TO_CREATE_RECOGNIZER: A$ = "Failed to create TextRecognizer"
  CASE %SAPIERR_FAILED_TO_CREATE_CONTEXT: A$ = "Failed to create RecognizerContext"
  CASE %SAPIERR_FAILED_TO_CREATE_GRAMMAR: A$ = "Failed to create grammar"
  CASE %SAPIERR_FAILED_TO_RESET_GRAMMAR: A$ = "Failed to reset grammar"
  CASE %SAPIERR_FAILED_TO_GET_RULE: A$ = "Failed to get rule"
  CASE %SAPIERR_FAILED_TO_SET_NOTIFY: A$ = "Failed to set notify event"
  CASE %SAPIERR_FAILED_TO_SET_INTEREST: A$ = "Failed to set interest"
  CASE %SAPIERR_FAILED_TO_ACTIVATE_RULE: A$ = "Failed to activate rule"
  CASE %SAPIERR_FAILED_TO_RESUME_CONTEXT: A$ = "Failed to resume context"
  CASE -1000 TO -99999: A$ = "Failed to add word #"+FORMAT$(-1000-ERRCODE&)+" to grammar"
  CASE ELSE: A$ = "Unknown error ("+FORMAT$(ERRCODE&)+")
  END SELECT

  SAPIGETERRORMSG$ = A$
END FUNCTION



'Alle verfügbaren Stimmen ermitteln
SUB GETVOICES(VNAME$(), VATTR$())
  LOCAL I&, N&, A$, B$
  LOCAL v AS ISpeechVoice
  LOCAL tokens AS ISpeechObjectTokens
  LOCAL t AS ISpeechObjectToken

  v = NEWCOM CLSID $CLSID_SPEECHLIB
  IF ISNOTHING(v) THEN EXIT SUB

  tokens = v.GetVoices()
  N& = tokens.Count
  REDIM VNAME$(N&-1), VATTR$(N&-1)
  FOR I& = 0 TO N&-1
    t = tokens.Item(I&)
    VNAME$(I&) = t.GetAttribute("Name")
    A$ = "?"
    B$ = UCASE$(t.GetAttribute("Gender"))
    IF B$ <> "" THEN A$ = LEFT$(B$, 1)
    B$ = UCASE$(t.GetAttribute("Age"))
    IF B$ <> "" THEN A$ = A$+LEFT$(B$, 1) ELSE A$ = A$+"?"
    A$ = A$+t.GetAttribute("Language")
    VATTR$(I&) = A$
  NEXT
END SUB



'Sprechgeschwindigkeit (-10 bis +10)
FUNCTION SAPISetRate(rate&, a$$) AS WSTRING
  rate& = MIN&(10, MAX&(-10, rate&))
  SAPISetRate = "<RATE SPEED="+CHR$(34)+FORMAT$(rate&)+CHR$(34)+">"+a$$+"</RATE>"
END FUNCTION



'Lautstärke (1 bis 100)
FUNCTION SAPISetVolume(volume&, a$$) AS WSTRING
  volume& = MIN&(100, MAX&(1, volume&))
  SAPISetVolume = "<VOLUME LEVEL="+CHR$(34)+FORMAT$(volume&)+CHR$(34)+">"+a$$+"</VOLUME>"
END FUNCTION



'Stimmhöhe (-10 bis +10)
FUNCTION SAPISetPitch(pitch&, a$$) AS WSTRING
  pitch& = MIN&(10, MAX&(-10, pitch&))
  SAPISetPitch = "<PITCH MIDDLE="+CHR$(34)+FORMAT$(pitch&)+CHR$(34)+">"+a$$+"</PITCH>"
END FUNCTION



'Stimme
'Gender: 1=Male , 2=Female
'Lang: 407=German
FUNCTION SAPISetVoice(a$$, gender&, lang&, vname$) AS WSTRING
  LOCAL x$$

  IF gender& > 0 OR lang& > 0 OR vname$ <> "" THEN
    x$$ = " REQUIRED="+CHR$(34)
    IF gender& > 0 THEN x$$ = x$$+"GENDER="+IIF$(gender& = 1, "MALE", "FEMALE")
    IF lang& > 0 THEN
      IF RIGHT$(x$$, 1) <> CHR$(34) THEN x$$ = x$$+";"
      x$$ = x$$+"LANGUAGE="+FORMAT$(lang&)
    END IF
    IF vname$ <> "" THEN x$$ = " REQUIRED="+CHR$(34)+"NAME="+vname$
    x$$ = x$$+CHR$(34)
  END IF
  SAPISetVoice = "<VOICE"+x$$+">"+a$$+"</VOICE>"
END FUNCTION



'Pause (in ms)
FUNCTION SAPISilence(millisecs&) AS WSTRING
  SAPISilence = "<SILENCE MSEC="+CHR$(34)+FORMAT$(millisecs&)+CHR$(34)+"/>"
END FUNCTION



'Satz betonen
FUNCTION SAPIEmph(a$$) AS WSTRING
  SAPIEmph = "<EMPH>"+a$$+"</EMPH>"
END FUNCTION



'Grammatik für Spracherkennung setzen
'Parameter:
'language = Sprache der Grammatik (0 = UserDefaultLanguage, sonst MAKELANGID(%LANG_ENGLISH, %SUBLANG_ENGLISH_US) oder MAKELANGID(%LANG_GERMAN, %SUBLANG_GERMAN))
'Rückgabewerte:
'1 = Erfolg
'0 = Grammatik konnte nicht erzeugt werden
'-1 .. -xxx =
FUNCTION SAPICREATEGRAMMAR&(BYVAL language&, grammarwords$())
  LOCAL hr&, i&, rulestate&, interest&&
  LOCAL wordtoadd AS WSTRINGZ*256, separator AS WSTRINGZ*2

  'Interface erzeugen
  hr& = pIRecoContext.CreateGrammar(%GRAMMAR_ID, pIRecoGrammar)
  IF hr& <> %S_OK THEN
    SAPICREATEGRAMMAR& = %SAPIERR_FAILED_TO_CREATE_GRAMMAR
    EXIT FUNCTION
  END IF

  'Sprache einstellen
  IF language& = 0 THEN language& = SpGetUserDefaultUILanguage
  hr& = pIRecoGrammar.ResetGrammar(language&)
  IF hr& <> %S_OK THEN
    SAPICREATEGRAMMAR& = %SAPIERR_FAILED_TO_RESET_GRAMMAR
    EXIT FUNCTION
  END IF
  hr& = pIRecoGrammar.GetRule($GRAMMAR_RULE, 0, %SPRAF_TopLevel OR %SPRAF_Active, %TRUE, rulestate&)
  IF hr& <> %S_OK THEN
    SAPICREATEGRAMMAR& = %SAPIERR_FAILED_TO_GET_RULE
    EXIT FUNCTION
  END IF

  'Wörter einstellen
  separator = " "
  FOR i& = 0 TO UBOUND(grammarwords$())
    wordtoadd = grammarwords$(i&)
    hr& = pIRecoGrammar.AddWordTransition(rulestate&, BYVAL %NULL, wordtoadd, separator, %SPWT_LEXICAL, 1, BYVAL %NULL)
    IF hr& <> %S_OK THEN
      SAPICREATEGRAMMAR& = -1000-i&
      EXIT FUNCTION
    END IF
  NEXT i&

  'Sprache abschließen
  pIRecoGrammar.Commit(0)

  'Grammatik aktivieren
  pIRecoContext.SetNotifyWin32Event
  STT_NOTIFICATION_HANDLE& = pIRecoContext.GetNotifyEventHandle
  IF STT_NOTIFICATION_HANDLE& = %INVALID_HANDLE_VALUE THEN
    SAPICREATEGRAMMAR& = %SAPIERR_FAILED_TO_SET_NOTIFY
    EXIT FUNCTION
  END IF
  interest&& = SPFEI(%SPEI_RECOGNITION)
  hr& = pIRecoContext.SetInterest(interest&&, interest&&)
  IF hr& <> %S_OK THEN
    SAPICREATEGRAMMAR& = %SAPIERR_FAILED_TO_SET_INTEREST
    EXIT FUNCTION
  END IF
  pIRecoGrammar.SetRuleState($GRAMMAR_RULE, 0, %SPRS_ACTIVE)
  IF hr& <> %S_OK THEN
    SAPICREATEGRAMMAR& = %SAPIERR_FAILED_TO_ACTIVATE_RULE
    EXIT FUNCTION
  END IF
  hr& = pIRecoContext.Resume(0)
  IF hr& <> %S_OK THEN
    SAPICREATEGRAMMAR& = %SAPIERR_FAILED_TO_RESUME_CONTEXT
    EXIT FUNCTION
  END IF

  SAPICREATEGRAMMAR& = 1
END FUNCTION



'Wartet bis der Benutzer etwas spricht
FUNCTION WAITUSERSPEAKINGUNICODE AS WSTRING
  LOCAL a$$, hr&, eventcount&, pText&, p&
  LOCAL h&()
  LOCAL e() AS SPEVENT
  LOCAL recognized AS ISpRecoResult
  DIM h&(0), e(%MAX_TEXTRECOGNITION_EVENTS-1)

  'auf Benutzereingabe warten
  h&(0) = STT_NOTIFICATION_HANDLE&
  WaitForMultipleObjects(1, h&(0), %FALSE, %INFINITE)

  'gesprochenen Text ermitteln
  hr& = pIRecoContext.GetEvents(%MAX_TEXTRECOGNITION_EVENTS, e(0), eventcount&)
  IF hr& <> %S_OK AND hr& <> %S_FALSE THEN
    EXIT FUNCTION
  END IF

  'DWORD Pointer in ISpRecoResult unwandeln
  POKE LONG, VARPTR(recognized), e(0).lParam

  'erkannten Text auslesen
  hr& = recognized.GetText(%SP_GETWHOLEPHRASE, %SP_GETWHOLEPHRASE, %FALSE, pText&, BYVAL %NULL)
  IF hr& = %S_OK THEN
    p& = pText&
    WHILE PEEK(INTEGER, p&) <> 0
      p& = p&+2
    WEND
    a$$ = PEEK$$(pText&, (p&-pText&)/2)
  END IF

  'Speicherplatz wieder freigeben
  CoTaskMemFree(pText&)

  WAITUSERSPEAKINGUNICODE = a$$
END FUNCTION



'SAPI initialisieren
'md& = 0 : Nur Text-To-Speech initialisieren
'md& = 1 : Text-To-Speech und Speech-To-Text initialisieren
FUNCTION SAPIINIT&(md&)
  LOCAL hr&

  'prüfen, ob Initialisierung bereits zuvor ausgeführt wurde
  IF SAPIINSTALLED& <> 0 THEN
    SAPIINIT& = MAX&(SAPIINSTALLED&, 0)
    EXIT FUNCTION
  END IF

  'Schnittstelle zur SAPI Bibliothek erzeugen
  pISpVoice = NEWCOM CLSID $CLSID_SPEECHLIB
  IF ISNOTHING(pISpVoice) THEN
    SAPIINSTALLED& = -1
    SAPIINIT& = %SAPIERR_FAILED_TO_CREATE_SPVOICE
    EXIT FUNCTION
  END IF
  '
  IF md& = 1 THEN
    pITextRecognizer = NEWCOM CLSID $CLSID_TOTEXTLIB
    IF ISNOTHING(pITextRecognizer) THEN
      SAPIINSTALLED& = -1
      SAPIINIT& = %SAPIERR_FAILED_TO_CREATE_RECOGNIZER
      EXIT FUNCTION
    ELSE
      hr& = pITextRecognizer.CreateRecoContext(pIRecoContext)
      IF hr& <> %S_OK THEN
        SAPIINSTALLED& = -1
        SAPIINIT& = %SAPIERR_FAILED_TO_CREATE_CONTEXT
        EXIT FUNCTION
      END IF
      pIRecoContext.Pause(0)
    END IF
  END IF

  SAPIINSTALLED& = 1
  SAPIINIT& = 1
END FUNCTION



'Beendet die Kommunikation mit der SAPI Bibliothek
SUB SAPIEND
  pISpVoice = NOTHING
  pITextRecognizer = NOTHING
  pIRecoContext = NOTHING
  SAPIINSTALLED& = 0
END SUB



'Stellt die Optionen zur Sprachausgabe ein
SUB SAPIOPTIONS(NOWAIT&)
  SAPINOWAIT& = NOWAIT&
END SUB



'Wartet bis die aktuelle Sprachausgabe abgeschlossen ist (oder das Timeout erreicht ist)
'Liefert TRUE zurück wenn die Sprachausgabe abgeschlossen wurde; FALSE wenn das Timeout erreicht wurde
FUNCTION SAPIWAITUNTILDONE&(TIMEOUTMS&)
  LOCAL R&

  IF ISOBJECT(pISpVoice) THEN R& = pISpVoice.WaitUntilDone(TIMEOUTMS&)
  SAPIWAITUNTILDONE& = IIF(R& = 0, -1, 0)
END FUNCTION



'Spricht einen Text
SUB SAPISpeak(a$$)
  LOCAL r&
  LOCAL ulStreamNumber AS DWORD
  LOCAL wszText AS WSTRINGZ*65536

  IF SAPIINSTALLED& = 0 THEN
    r& = SAPIINIT&(0)
    IF r& = 0 THEN EXIT SUB
  END IF
  IF SAPIINSTALLED& < 0 THEN EXIT SUB

  wszText = a$$
  pISpVoice.Speak(wszText, %SPF_ASYNC OR %SPF_PURGEBEFORESPEAK, ulStreamNumber)
  IF SAPINOWAIT& = 0 THEN pISpVoice.WaitUntilDone(&HFFFFFFFF)
END SUB
