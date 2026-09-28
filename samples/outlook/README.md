# Ekte Outlook-filer (.msg) til testing

Filene er testfiler fra to åpne prosjekter, brukt under MIT-lisens (se lisensfilene).
`.msg` kan ikke lages med biblioteket vi bruker (extract-msg), så ekte filer er
eneste måte å teste mot Outlooks faktiske format.

| Fil | Kilde | Tester |
|---|---|---|
| `EmailWithInnerMailAndAttachments.msg` | [MSGReader](https://github.com/Sicos1977/MSGReader) `MsgReaderTests/SampleFiles/` | e-post i e-post, PDF-vedlegg på to nivåer |
| `EmailWithSpecialCharsInSubject_2.msg` | MSGReader | fransk emne (tegnkoding) |
| `RtfWithShortRussianString.msg` | MSGReader | russisk tekst |
| `kontakt-Swetlana.msg` (opprinnelig `Swetlana.msg`) | [mapi](https://github.com/hfig/mapi) `tests/_files/` | Outlook-kontakt, ikke e-post |
| `mapi-sample.msg` (opprinnelig `sample.msg`) | mapi | HTML-e-post med tekstvedlegg |

Lisenser: `LICENSE-MSGReader.txt` (MIT), `LICENSE-mapi.txt` (MIT, Copyright (c) 2018 hfig).
