{runModule, VBRuntimeError} = require '../vb'
assert = require 'assert'

# from https://gist.github.com/734620
repr = (o, depth=0, max=2) ->
  if depth > max
    '<..>'
  else
    switch typeof o
      when 'string' then "\"#{o.replace /"/g, '\\"'}\""
      when 'function' then 'function'
      when 'object'
        if o is null then 'null'
        if _.isArray o
          '[' + [''+repr(e, depth + 1, max) for e in o] + ']'
        else
          '{' + [''+k+':'+repr(o[k], depth + 1, max) for k in _.keys(o)] + '}'
      when 'undefined' then 'undefined'
      else o

run = (code, expected) ->
    log = ''
    code = """Function Foo()
              #{code}
              End Function"""
    #console.log "code: `#{code}`"
    module = runModule code, DoCmd:
                            dot: (name) ->
                                (args...) ->
                                    spec = (repr a for a in args).join ','
                                    log += "#{name}(#{spec})\n"
    if expected?
        module.Foo()
        assert.strictEqual log, expected
    module

assert_js = (module, expected_obj) ->
    for fn, expected of expected_obj
        match = module[fn].toString().match /^function \(\) \{\s*([\s\S]*)\s*\}$/
        body = match[1].replace /\s*\n\s*/g, '\n' # eat whitespaces around \n
        assert.strictEqual body, expected

test_foo_close = ({before, after, after_spec, before_func}) ->
    fill = (s) -> if s? then s + '\n' else ''

    before_func = fill before_func
    before = fill before
    after = fill after
    unless after_spec?
        after_spec = ''

    code = """#{before_func}Function Foo() #{after_spec}
                #{before}DoCmd.Close
                #{after}End Function"""

    assert_js runModule(code), Foo: "ns('DoCmd').dot('Close')();\n"

suite 'Modules -', ->
    test 'empty', ->
        run '', ''

    test 'one line', ->
        run 'DoCmd.Close', 'Close()\n'

    test 'bracketed', ->
        m = runModule """Function Foo()
                           Me.Parent![Customer Orders].Requery
                           DoCmd.Close
                         End Function"""
        assert_js m,
            Foo: """ns('Me').dot('Parent').bang('Customer Orders').dot('Requery')();
                    ns('DoCmd').dot('Close')();\n"""

    # SKIPPED
    if 0 then test 'function().property', ->
        m = runModule """CurrentDb().Properties("StartupForm")"""
        assert_js m, Foo: "ns('CurrentDb')().dot('Properties')('StartupForm');"

    test 'nested dot', ->
        assert_js run('DoCmd.Nested.Close'),
                  Foo: "ns('DoCmd').dot('Nested').dot('Close')();\n"

    test 'arguments', ->
        run 'DoCmd.OpenForm "Main Switchboard", 123',
            'OpenForm("Main Switchboard",123)\n'

    test 'arguments missing', ->
        run 'DoCmd.OpenReport "Sales by Category", 1, , 3',
            'OpenReport("Sales by Category",1,undefined,3)\n'

    test 'one argument', ->
        run 'DoCmd.OpenForm "Main Switchboard"',
            'OpenForm("Main Switchboard")\n'

    test 'one braced argument', ->
        run 'DoCmd.OpenForm ("Main Switchboard")',
            'OpenForm("Main Switchboard")\n'

    test 'two braced arguments', ->
        run 'DoCmd.MoveSize (1), (2)', 'MoveSize(1,2)\n'

    test 'dotted argument', ->
        assert_js run('MsgBox A.B'),
                  Foo: "ns('MsgBox')(ns('A').dot('B'));\n"

    test 'numbers', ->
        run 'DoCmd.Foo 1, 23, 456',
            'Foo(1,23,456)\n'

    test 'comment', ->
        run "' hi there!", ''

    test 'double comment', ->
        run "' hi there!\n' bye!", ''

    test 'double comment indented', ->
        run "  ' hi there!\n  ' bye!", ''

    test 'end comment', ->
        run "DoCmd.Close ' this is comment", 'Close()\n'

    test 'complex indented', ->
        run "  ' Closes Startup form.\n" +
            "  ' Used in OnClick property...\n" +
            "   DoCmd.Close\n" +
            "   DoCmd.OpenForm (\"Main Switchboard\")",
            'Close()\nOpenForm("Main Switchboard")\n'

    test 'strange', ->
        run 'DoCmd.EndFunction', 'EndFunction()\n'

    test 'leading empty line', ->
        m = runModule "\nFunction Foo()\nDoCmd.Close\nEnd Function"
        assert_js m, Foo: "ns('DoCmd').dot('Close')();\n"

    test 'empty module', ->
        runModule ''

    test 'declarations stub', ->
        test_foo_close before_func: """Option Compare Database
                                       Option Explicit
                                       Attribute VB_Exposed = False
                                       Dim path As String
                                       ' Decrarations end here"""

    test 'func_def arguments stub', ->
        m = runModule """Sub Foo(A, B As Integer, ByVal C As Integer)
                           DoCmd.Close
                         End Sub"""
        assert_js m, Foo: "ns('DoCmd').dot('Close')();\n"

    test 'function As stub', ->
        test_foo_close after_spec: 'As Boolean'

    test 'function Error Resume stub', ->
        test_foo_close before: 'On Error Resume Next'

    test 'function Error GoTo stub', ->
        test_foo_close before: 'On Error GoTo LabelName'

    test 'Resume stub', ->
        test_foo_close
            before: 'On Error GoTo LabelName'
            after: 'Resume LabelName'

    test 'Exit Function', ->
        run """DoCmd.Close
               Exit Function
               DoCmd.WillNotRun""", 'Close()\n'

    test 'Label stub', ->
        run """On Error GoTo FooError
                   DoCmd.Close
               FooExit:
                   Exit Function
               FooError:
                   DoCmd.HandleError
                   Resume FooExit""", 'Close()\n'

    test 'With stub', ->
      test_foo_close after: """With Application.FileDialog(msoFileDialogFilePicker)
                                 .Title = "Name"
                                 .Filters.Add "All Files", "*.*"
                               End With"""

    test 'plain GoTo stub', ->
      test_foo_close after: 'GoTo LabelName'

    test 'Const Set stub', ->
        test_foo_close before: """Const someConst = 42
                                  Set someVal = 43"""

    test 'Dim stub', ->
        test_foo_close before: """Dim returnValue As Boolean, path As String
                                  Static db As DAO.Database
                                  Dim Num
                                  """

    test 'assign stub', ->
        test_foo_close before: """foo = "Foo"
                                  bar = False"""

    test 'If stub', ->
        test_foo_close after: """FooErr:
                                 If Err = 3270 Then
                                     DoCmd.Bla
                                     Resume FooExit
                                 End If"""

    test 'Like stub', -> # TODO move to test/expressions.coffee
        test_foo_close before: """If Not Foo Like "[A-Z]" Then
                                     DoCmd.Bla
                                  End If"""

    test 'If Not EndIf stub', ->
        test_foo_close before: """If Not IsNull(Me!Photo) Then
                                      hideImageFrame
                                  EndIf"""

    test 'If ElseIf Else stub', ->
        test_foo_close before: """If IsItReplica() Then
                                      DoCmd.TellAboutIt
                                  ElseIf IsLoaded("Product List") Then
                                      DoCmd.OpenForm strDocName
                                  Else
                                      DoCmd.DoSomething
                                      DoCmd.DoSomethingElse
                                  End If"""

    test 'If single line stub', ->
        test_foo_close before: "If IsItReplica() Then DoCmd.TellAboutIt"


    test 'Condition in braces stub', ->
        test_foo_close before: """If (IsItReplica()) Then
                                      DoCmd.TellAboutIt
                                  End If"""

    test 'If Or _ stub', ->
        test_foo_close before: """
            If (CurrentDb().Properties("StartupForm") = "Startup" Or _
                CurrentDb().Properties("StartupForm") = "Form.Startup") Then
                Forms!Startup!HideStartupForm = False
            End If"""

    test 'If Or And stub', ->
        test_foo_close before: """
            If (Aa = "") Or (Ba < 1) _
               Or (C <> D) And (D >= 5) Then
                E = 0
            End If"""

    test 'If Or CrLf', ->
        test_foo_close before: 'If A _\r\nOr B Then\r\nC = 0\r\nEnd If'

    test 'Select Case stub', ->
        test_foo_close before: """
          Select Case Me!ReportToPrint
            Case 1, 3 To 5
              DoCmd.OpenReport "Sales Totals by Amount", PrintMode
              DoCmd.Quit
            Case Is >< 10
              DoCmd.Quit
            Case Else
              DoCmd.Quit
          End Select"""

    test 'several functions', ->
        m = runModule """Function Foo()
                           DoCmd.Open
                         End Function
                         ' Second function
                         Private Sub Bar()
                           DoCmd.Close
                         End Sub"""
        assert_js m,
            Foo: "ns('DoCmd').dot('Open')();\n"
            Bar: "ns('DoCmd').dot('Close')();\n"
