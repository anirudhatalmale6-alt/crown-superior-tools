Attribute VB_Name = "CrownUnited"
Option Explicit

' =====================================================================
' Crown Superior - the two United pages between Coverages and the MVR
' ---------------------------------------------------------------------
' 1. "Additional Driver(s)/Vehicle(s) have been found in the household"
'    (Cmn_Lightspeed.mac). United lists people and cars it found at the
'    address. Every driver gets Unknown and every vehicle gets Do not add
'    - however many there are - then Continue.
'
' 2. Quote Summary and Payment Options (cmn_Premium.mac), done the way
'    Tony does it by hand:
'      - OK the pop-ups (previous policy balance, agency fees)
'      - Select Payment Plan under 16.67% DOWN PAY, OK the pop-up again
'      - down payment + the red Previous Policy Balance (it can be a
'        credit, so it can take money off) = what goes in Cash
'      - tick Credit/Debit Card, top right, OK the agency fees pop-up
'    Then his macro presses Continue as before.
'
'    These are NOT the rates - the MVR has not run yet. The price the
'    customer gets is still read from the payment page after the MVR.
'
' Each step is written on the QuoteLook sheet, so a run that stops says
' where and why.
' =====================================================================

' What the payment page said, kept for the QuoteLook line.
Public CrownUnitedDown As String
Public CrownUnitedPrevBalance As String
Public CrownUnitedPayToday As String

Private Const LS_WAIT_SECONDS As Long = 45
Private Const PREMIUM_WAIT_SECONDS As Long = 30


' ---------------------------------------------------------------------
' 1. Extra drivers / vehicles
'
' Call right after Continue on the Coverages page. It waits for whichever
' comes next. If it is the extra drivers/vehicles page it answers it -
' as many times as United shows it - and returns once the payment page is
' there. If neither page turns up it writes down what is on screen and
' returns True anyway, so his own wait reports it as before.
'
' False only when United raised an alert that means stop.
' ---------------------------------------------------------------------
Public Function CrownUnitedLightspeed(ByVal drv As Object) As Boolean
    Dim started As Date
    Dim answer As String
    Dim rounds As Long

    started = Now

    Do
        If Not CrownHandleAlert(drv) Then Exit Function

        answer = CrownUnitedJs(drv, CrownLightspeedJs())

        If answer = "premium" Then
            CrownUnitedLightspeed = True

            Exit Function
        End If

        If Left$(answer, 4) = "done" Then
            rounds = rounds + 1
            CrownNoteBlocker "United extra drivers/vehicles page: " & Mid$(answer, 6)

            If rounds > 3 Then
                CrownNotePage drv, "United: the extra drivers/vehicles page kept coming back"
                CrownLastAlert = "United: the extra drivers/vehicles page kept coming back"

                Exit Function
            End If

            drv.Wait 3000
            started = Now
        ElseIf answer = "no continue" Then
            CrownNotePage drv, "United: extra drivers/vehicles page answered, but no Continue button found"
            CrownLastAlert = "United: no Continue button on the extra drivers/vehicles page"

            Exit Function
        End If

        drv.Wait 1000
        DoEvents
    Loop While (Now - started) * 86400 < LS_WAIT_SECONDS

    CrownNotePage drv, "United: after Coverages neither the extra drivers page nor the payment page came up"
    CrownUnitedLightspeed = True
End Function


' Every row with a choice of radios gets its LAST one: Unknown for a
' driver (Add / Exclude / Unknown), Do not add for a vehicle (Add / Do
' not add). Done by row, not by a fixed count, so ten extra drivers are
' handled the same as one. Then Continue.
Private Function CrownLightspeedJs() As String
    CrownLightspeedJs = _
        "var txt=document.body?(document.body.innerText||''):'';" & _
        "var ls=/lightspeed/i.test(location.href)||/additional driver\(s\)\/vehicle\(s\) have been found/i.test(txt);" & _
        "if(!ls){if(document.getElementById('rowPlan')||document.getElementById('btnPaymentPlanDBI6'))return 'premium';return 'wait';}" & _
        "if(document.readyState!=='complete')return 'wait';" & _
        "var picks=[],d=0,v=0,rows=document.querySelectorAll('tr');" & _
        "for(var i=0;i<rows.length;i++){var r=rows[i],rad=[],all=r.querySelectorAll('input[type=radio]');" & _
        "for(var k=0;k<all.length;k++){if(all[k].closest('tr')===r)rad.push(all[k]);}" & _
        "if(rad.length<2)continue;picks.push(rad[rad.length-1]);if(rad.length>=3)d++;else v++;}" & _
        "if(!picks.length){var g={},order=[],rs=document.querySelectorAll('input[type=radio]');" & _
        "for(var m=0;m<rs.length;m++){var n=rs[m].name||('#'+m);if(!g[n]){g[n]=[];order.push(n);}g[n].push(rs[m]);}" & _
        "for(var o=0;o<order.length;o++){var grp=g[order[o]];if(grp.length<2)continue;picks.push(grp[grp.length-1]);if(grp.length>=3)d++;else v++;}}" & _
        "var go=null,bs=document.querySelectorAll('input[type=button],input[type=submit],button,a');" & _
        "for(var b=0;b<bs.length;b++){var t=((bs[b].value||'')+' '+(bs[b].innerText||'')).replace(/\s+/g,' ').trim().toLowerCase();" & _
        "if(t==='continue'){go=bs[b];break;}}" & _
        "if(!go)return 'no continue';" & _
        "setTimeout(function(){for(var p=0;p<picks.length;p++){if(!picks[p].checked)picks[p].click();}" & _
        "setTimeout(function(){go.click();},700);},50);" & _
        "return 'done '+d+' driver(s) set to Unknown, '+v+' vehicle(s) set to Do not add';"
End Function


' ---------------------------------------------------------------------
' 2. The payment page (first time, before the MVR)
'
' True when the 16.67% plan is chosen, Cash holds down + previous
' balance and Credit/Debit Card is ticked. False with the reason in
' CrownLastAlert (and the page on QuoteLook) otherwise.
' ---------------------------------------------------------------------
Public Function CrownUnitedPremium(ByVal drv As Object) As Boolean
    Dim answer As String, figures As String
    Dim down As Double, balance As Double, total As Double
    Dim cash As String
    Dim started As Date
    Dim tries As Long

    CrownUnitedDown = ""
    CrownUnitedPrevBalance = ""
    CrownUnitedPayToday = ""

    If Not CrownUnitedSettle(drv) Then Exit Function

    ' --- 16.67% DOWN PAY -------------------------------------------------
    For tries = 1 To 3
        answer = CrownUnitedJs(drv, CrownPlanJs())

        If answer = "already" Then Exit For

        If answer = "clicked" Then
            drv.Wait 2500
            If Not CrownUnitedSettle(drv) Then Exit Function
        ElseIf answer = "no button" Or answer = "no 16.67 column" Then
            ' Not where it was looked for - his own button, as before.
            CrownNoteBlocker "United payment page: " & answer & " - trying btnPaymentPlanDBI6"
            On Error Resume Next
            drv.FindElementById("btnPaymentPlanDBI6").Click
            On Error GoTo 0
            drv.Wait 2500
            If Not CrownUnitedSettle(drv) Then Exit Function
        Else
            drv.Wait 1500
        End If
    Next tries

    If CrownUnitedJs(drv, CrownPlanJs()) <> "already" Then
        CrownNotePage drv, "United payment page: could not choose 16.67% DOWN PAY"
        CrownLastAlert = "United: could not choose the 16.67% plan on the payment page"

        Exit Function
    End If

    ' --- down + previous balance ------------------------------------------
    figures = CrownUnitedJs(drv, CrownFiguresJs())

    If InStr(1, figures, "|") = 0 Then
        CrownNotePage drv, "United payment page: could not read the down payment (" & figures & ")"
        CrownLastAlert = "United: could not read the 16.67% down payment"

        Exit Function
    End If

    down = Val(Split(figures, "|")(0))
    balance = Val(Split(figures, "|")(1))
    total = Round(down + balance, 2)

    CrownUnitedDown = CrownMoney(down)
    CrownUnitedPrevBalance = CrownMoney(balance)
    CrownUnitedPayToday = CrownMoney(total)

    ' --- Insured Pays Agency, then Cash -----------------------------------
    ' United starts the Payment Method box on "Credit Card (UAIC collects
    ' from Insured)" with the whole amount in it. Cash typed under that is
    ' a second payment of the same amount, and Continue answers "Amount
    ' for method of payment should be equal to total amount due" - which
    ' stopped every United quote from 15 to 17 Oct. Tony picks Insured
    ' Pays Agency first, then types Cash.
    answer = CrownUnitedJs(drv, CrownAgencyMethodJs())

    If Left$(answer, 2) <> "ok" Then
        CrownNotePage drv, "United payment page: " & answer
        CrownLastAlert = "United: could not choose Insured Pays Agency on the payment page"

        Exit Function
    End If

    CrownNoteBlocker "United payment method: " & answer
    drv.Wait 1200
    If Not CrownHandleAlert(drv) Then Exit Function

    cash = CrownUnitedTypeCash(drv, CrownUnitedPayToday)

    If Left$(cash, 2) <> "ok" Then
        CrownNotePage drv, "United payment page: " & cash
        CrownLastAlert = "United: could not find the Cash box on the payment page"

        Exit Function
    End If

    drv.Wait 800
    If Not CrownHandleAlert(drv) Then Exit Function

    ' --- Credit/Debit Card, top right --------------------------------------
    answer = CrownUnitedJs(drv, CrownCardJs())

    If Left$(answer, 2) <> "ok" Then
        CrownNotePage drv, "United payment page: " & answer
        CrownLastAlert = "United: could not tick Credit/Debit Card on the payment page"

        Exit Function
    End If

    ' "Agency fees can not be included in credit card payments", and the
    ' page may redraw after it.
    drv.Wait 1500
    If Not CrownUnitedSettle(drv) Then Exit Function

    ' A redraw can put Cash back to what United thinks. Put ours back.
    started = Now
    Do
        cash = CrownUnitedJs(drv, CrownCashReadJs())
        If Abs(Val(cash) - total) < 0.005 Then Exit Do

        CrownUnitedJs drv, CrownAgencyMethodJs()
        drv.Wait 600
        If Not CrownHandleAlert(drv) Then Exit Function
        CrownUnitedTypeCash drv, CrownUnitedPayToday
        drv.Wait 800
        If Not CrownHandleAlert(drv) Then Exit Function
    Loop While (Now - started) * 86400 < 10

    If Abs(Val(cash) - total) >= 0.005 Then
        CrownNotePage drv, "United payment page: Cash shows " & cash & " but it should be " & CrownUnitedPayToday
        CrownLastAlert = "United: Cash would not stay at " & CrownUnitedPayToday

        Exit Function
    End If

    ' The agency-fees pop-up can come a few seconds after the tick - after
    ' everything above has already looked. One left open is a page nobody
    ' can press Continue on (Tony's screenshot of 17 Oct),
    ' so watch for it for a few seconds more before handing back.
    started = Now
    Do
        If Not CrownHandleAlert(drv) Then Exit Function
        drv.Wait 500
        DoEvents
    Loop While (Now - started) * 86400 < 4

    ' Every amount box on the page, hidden ones too - if United still says
    ' the payment does not match its total, this line shows what it wanted.
    CrownNoteBlocker "United payment page amounts: " & CrownUnitedJs(drv, CrownAmountsJs())

    ' Every choice on the page - each tick box and round button with its
    ' own words and whether it is on, and every visible box with what is
    ' in it. United keeps answering "Amount for method of payment should be
    ' equal to total amount due" although every amount above matches, so
    ' the method chosen is the likely problem, and this shows which one.
    CrownNoteBlocker "United payment page choices: " & CrownUnitedJs(drv, CrownChoicesJs())

    CrownNoteBlocker "United payment page: 16.67% plan, down " & CrownUnitedDown & _
                     " + previous balance " & CrownUnitedPrevBalance & _
                     " = " & CrownUnitedPayToday & " in Cash, Credit/Debit Card ticked (" & answer & ")"

    CrownUnitedPremium = True
End Function


' OK every pop-up, then wait for the payment table to be back on screen.
' False only when United said stop.
Private Function CrownUnitedSettle(ByVal drv As Object) As Boolean
    Dim started As Date
    Dim answer As String

    started = Now

    Do
        If Not CrownHandleAlert(drv) Then Exit Function

        answer = CrownUnitedJs(drv, _
            "return (document.readyState==='complete'&&document.getElementById('rowPlan'))?'ready':'wait';")

        If answer = "ready" Then
            ' One more look: the previous-balance pop-up comes a moment
            ' after the page itself.
            drv.Wait 700
            If Not CrownHandleAlert(drv) Then Exit Function

            CrownUnitedSettle = True

            Exit Function
        End If

        drv.Wait 700
        DoEvents
    Loop While (Now - started) * 86400 < PREMIUM_WAIT_SECONDS

    CrownNotePage drv, "United: the payment page did not come back"
    CrownLastAlert = "United: the payment page did not come back"
End Function


' The 16.67% column, found by its heading like his WritePaymentPlanFull
' does, and its Select Payment Plan button pressed.
Private Function CrownPlanJs() As String
    CrownPlanJs = _
        "var pr=document.getElementById('rowPlan');if(!pr)return 'no plan table';" & _
        "var col=-1;for(var i=0;i<pr.cells.length;i++){if(/16\.67\s*%/.test(pr.cells[i].innerText||'')){col=i;break;}}" & _
        "if(col<0)return 'no 16.67 column';" & _
        "var tb=pr.closest('table'),rs=tb.rows,btn=null,sel=false;" & _
        "for(var j=0;j<rs.length;j++){var c=rs[j].cells[col];if(!c)continue;" & _
        "if(/payment plan selected/i.test(c.innerText||''))sel=true;" & _
        "var els=c.querySelectorAll('input,button,a');" & _
        "for(var k=0;k<els.length;k++){var t=((els[k].value||'')+' '+(els[k].innerText||'')).toLowerCase();" & _
        "if(t.indexOf('payment plan selected')>=0)sel=true;else if(t.indexOf('select payment plan')>=0)btn=els[k];}}" & _
        "if(sel)return 'already';if(!btn)return 'no button';" & _
        "setTimeout(function(){btn.click();},50);return 'clicked';"
End Function


' "down|previous balance". The balance is red, beside Payment Plan
' Options, and can be a credit: shown with a minus, in brackets or as CR.
' No balance shown means none.
Private Function CrownFiguresJs() As String
    CrownFiguresJs = _
        "function num(s){s=((s||'')+'').replace(/,/g,'');var m=s.match(/\d*\.?\d+/);if(!m)return null;" & _
        "var v=parseFloat(m[0]);return /-|\(|\bcr\b/i.test(s)?-v:v;}" & _
        "var pr=document.getElementById('rowPlan'),rd=document.getElementById('rowDown');" & _
        "if(!pr||!rd)return 'no plan table';" & _
        "var col=-1;for(var i=0;i<pr.cells.length;i++){if(/16\.67\s*%/.test(pr.cells[i].innerText||'')){col=i;break;}}" & _
        "if(col<0||!rd.cells[col])return 'no 16.67 column';" & _
        "var d=num(rd.cells[col].innerText);if(d===null)return 'no down payment';" & _
        "var t=document.body.innerText||'',m=t.match(/previous policy balance\s*:?\s*([^\n\r]*)/i);" & _
        "var b=m?num(m[1]):0;if(b===null)b=0;" & _
        "return d.toFixed(2)+'|'+b.toFixed(2);"
End Function


' The Cash box under Insured Pays Agency, found by its label.
Private Function CrownCashJs() As String
    CrownCashJs = _
        "var tds=document.querySelectorAll('td,th'),inp=null;" & _
        "for(var i=0;i<tds.length;i++){if((tds[i].innerText||'').trim().toLowerCase()!=='cash')continue;" & _
        "var tr=tds[i].parentNode;inp=tr.querySelector('input[type=text],input:not([type])');if(inp)break;}" & _
        "if(!inp)return 'no Cash box';" & _
        "if(inp.id&&arguments[0]==='?id')return 'id:'+inp.id;" & _
        "inp.focus();inp.value=arguments[0];" & _
        "['input','keyup','change'].forEach(function(n){inp.dispatchEvent(new Event(n,{bubbles:true}));});" & _
        "inp.blur();return 'ok '+inp.value;"
End Function

Private Function CrownCashReadJs() As String
    CrownCashReadJs = _
        "var tds=document.querySelectorAll('td,th');" & _
        "for(var i=0;i<tds.length;i++){if((tds[i].innerText||'').trim().toLowerCase()!=='cash')continue;" & _
        "var inp=tds[i].parentNode.querySelector('input[type=text],input:not([type])');" & _
        "if(inp)return (inp.value||'').replace(/,/g,'');}return '';"
End Function


' Cash typed the way a person types it - clear, type, Tab out - so every
' handler United has on that box runs (a value set by script alone can
' leave United's own running total behind). Script fill only if the box
' has no id to find it by.
Private Function CrownUnitedTypeCash(ByVal drv As Object, ByVal amount As String) As String
    Dim found As String
    Dim box As Object
    Dim k As New Selenium.Keys

    found = CrownUnitedJs(drv, CrownCashJs(), "?id")

    If Left$(found, 3) <> "id:" Then
        CrownUnitedTypeCash = CrownUnitedJs(drv, CrownCashJs(), amount)

        Exit Function
    End If

    On Error Resume Next
    Set box = drv.FindElementById(Mid$(found, 4))
    box.Click
    box.Clear
    box.SendKeys k.Control & "a"
    box.SendKeys amount
    box.SendKeys k.Tab

    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        CrownUnitedTypeCash = CrownUnitedJs(drv, CrownCashJs(), amount)

        Exit Function
    End If
    On Error GoTo 0

    drv.Wait 600
    CrownUnitedTypeCash = "ok " & CrownUnitedJs(drv, CrownCashReadJs())
End Function


' The round button for "Insured Pays Agency" in the Payment Method box:
' the one whose own part of the page - the largest piece around it that
' holds no other round button - says Insured Pays Agency.
Private Function CrownAgencyMethodJs() As String
    CrownAgencyMethodJs = _
        "var rs=document.querySelectorAll('input[type=radio]');" & _
        "for(var i=0;i<rs.length;i++){var r=rs[i],own=r.parentNode,p=own;" & _
        "while(p&&p.parentNode&&p.parentNode.querySelectorAll('input[type=radio]').length===1)p=p.parentNode;" & _
        "own=p||own;var t=((own.innerText||'')+' '+(r.nextSibling&&r.nextSibling.textContent||'')).replace(/\s+/g,' ');" & _
        "if(!/insured pays agency/i.test(t))continue;" & _
        "var how=' #'+(r.id||r.name||'');if(r.checked)return 'ok Insured Pays Agency already chosen'+how;" & _
        "r.click();return (r.checked?'ok Insured Pays Agency chosen':'ok clicked Insured Pays Agency')+how;}" & _
        "return 'no Insured Pays Agency choice';"
End Function


' "id[x] words" for every tick box / round button, "id=value words" for
' every visible text box, in page order.
Private Function CrownChoicesJs() As String
    CrownChoicesJs = _
        "function c(s){return ((s||'')+'').replace(/\s+/g,' ').trim();}" & _
        "function words(e){var l=e.id?document.querySelector('label[for=\''+e.id+'\']'):null;" & _
        "if(l)return c(l.innerText).substring(0,50);var p=e.parentNode,t='';" & _
        "for(var k=0;k<3&&p&&!t;k++){t=c(p.innerText);p=p.parentNode;}return t.substring(0,50);}" & _
        "var out=[],es=document.querySelectorAll('input');" & _
        "for(var i=0;i<es.length;i++){var e=es[i],ty=(e.type||'text').toLowerCase();" & _
        "if(ty==='hidden'||ty==='button'||ty==='submit')continue;" & _
        "if(ty==='radio'||ty==='checkbox'){out.push((e.id||e.name)+(e.checked?'[x]':'[ ]')+(e.disabled?'(off)':'')+' '+words(e));}" & _
        "else if(e.getClientRects().length){out.push((e.id||e.name)+'='+c(e.value)+(e.disabled||e.readOnly?'(off)':'')+' '+words(e));}}" & _
        "return out.join(' ; ').substring(0,4000);"
End Function


' id=value for every box whose name mentions an amount, a total, a balance,
' a payment or what is due - hidden ones included.
Private Function CrownAmountsJs() As String
    CrownAmountsJs = _
        "var out=[],ins=document.querySelectorAll('input');" & _
        "for(var i=0;i<ins.length;i++){var e=ins[i],n=(e.id||e.name||'');" & _
        "if(!/amt|amount|due|tot|cash|bal|down|pay|prem|fee/i.test(n))continue;" & _
        "if(e.type==='radio'||e.type==='checkbox'){if(e.checked)out.push(n+'=ticked');continue;}" & _
        "if(e.type==='button'||e.type==='submit')continue;" & _
        "out.push(n+'='+(e.value||'')+(e.type==='hidden'?'(h)':''));}" & _
        "return out.join(' ; ').substring(0,1500);"
End Function


' Credit/Debit Card at the top right, beside Bank Transfer(EFT). Found as
' the smallest cell holding both labels - the Credit Card further down
' (UAIC collects from Insured) is a different thing and must not be
' touched. His old code's rdPayOnline only if that cell is not found.
Private Function CrownCardJs() As String
    CrownCardJs = _
        "var best=null,tds=document.querySelectorAll('td');" & _
        "for(var i=0;i<tds.length;i++){var t=tds[i].innerText||'';" & _
        "if(!/credit\/debit card/i.test(t)||!/bank transfer/i.test(t))continue;" & _
        "if(!tds[i].querySelector('input[type=radio]'))continue;" & _
        "if(!best||t.length<(best.innerText||'').length)best=tds[i];}" & _
        "var r=best?best.querySelector('input[type=radio]'):document.getElementById('rdPayOnline');" & _
        "if(!r)return 'no Credit/Debit Card choice';" & _
        "var how=(best?'found by its label':'rdPayOnline')+' #'+(r.id||r.name||'');" & _
        "if(r.checked)return 'ok already ticked, '+how;" & _
        "setTimeout(function(){r.click();},50);return 'ok '+how;"
End Function


' ---------------------------------------------------------------------
' Previous insurance company (the first United page)
'
' His macro picks the company only when the customer's answer is spelled
' exactly like United's list; anything else leaves it unpicked and United
' says "Please select previous company name". Tony: "it can select other
' non standard". So: the same name if it is there (or close), otherwise
' Other Non Standard.
' ---------------------------------------------------------------------
Public Sub CrownUnitedPriorCompany(ByVal drv As Object, ByVal company As String)
    Dim answer As String

    answer = CrownUnitedJs(drv, _
        "function f(s){return (s||'').replace(/[^a-z0-9]+/gi,' ').trim().toLowerCase();}" & _
        "var s=document.getElementById('txtPrevCompName');if(!s)return 'no company list';" & _
        "var w=f(arguments[0]),pick=-1,how='';" & _
        "for(var i=0;i<s.options.length&&pick<0;i++){if(w&&f(s.options[i].text)===w){pick=i;how='same name';}}" & _
        "for(var i=0;i<s.options.length&&pick<0;i++){var t=f(s.options[i].text);" & _
        "if(w.length>=4&&t.length>=4&&(t.indexOf(w)>=0||w.indexOf(t)>=0)){pick=i;how='close match';}}" & _
        "var tries=[/other.*non.?standard/i,/non.?standard/i,/^other/i];" & _
        "for(var k=0;k<tries.length&&pick<0;k++){for(var i=0;i<s.options.length;i++){" & _
        "if(tries[k].test(s.options[i].text)){pick=i;how='not on the list - Other';break;}}}" & _
        "if(pick<0)return 'nothing to pick for '+arguments[0];" & _
        "s.selectedIndex=pick;setTimeout(function(){s.dispatchEvent(new Event('change',{bubbles:true}));},50);" & _
        "return 'chose '+s.options[pick].text.trim()+' ('+how+')';", _
        IIf(Len(Trim$(company)) = 0, " ", company))

    drv.Wait 800
    CrownNoteBlocker "United previous company '" & company & "': " & answer
End Sub


' Run a script and hand back its answer as text; an open alert or a page
' mid-load comes back as "wait" rather than as an error.
Private Function CrownUnitedJs(ByVal drv As Object, ByVal script As String, _
                               Optional ByVal arg As String = "") As String
    Dim answer As Variant

    On Error Resume Next
    Err.Clear

    If Len(arg) > 0 Then
        answer = drv.ExecuteScript(script, arg)
    Else
        answer = drv.ExecuteScript(script)
    End If

    If Err.Number <> 0 Then answer = "wait"
    Err.Clear
    On Error GoTo 0

    If IsNull(answer) Or IsEmpty(answer) Then answer = ""

    CrownUnitedJs = CStr(answer)
End Function


' 149.3 -> "149.30", whatever the PC's own number format.
Private Function CrownMoney(ByVal value As Double) As String
    CrownMoney = Replace(Format$(value, "0.00"), ",", ".")
End Function
