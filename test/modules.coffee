{runModule, compileModule, evaluate, VBRuntimeError} = require '../vb'
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
    log = []
    code = """Function Foo()
              #{code}
              End Function"""
    #console.log "code: `#{code}`"
    module = runModule code, DoCmd:
                            dot: (name) ->
                                (args...) ->
                                    spec = (repr a for a in args).join ','
                                    log.push "#{name}(#{spec})"
    if expected?
        module.Foo()
        assert.strictEqual log.join('\n'), expected
    module

assert_js = (vba, expected_obj) ->
  actual = evaluate compileModule vba
  for fn, expected of expected_obj
    match = actual[fn].toString().match /^function \(\) \{\s*([\s\S]*)\s*\}$/
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

    assert_js code, Foo: "ns('DoCmd').dot('Close')();\n"

foo = (code) -> "Function Foo()\n  #{code}\nEnd Function"

suite 'Modules -', ->
    test 'empty', ->
        run '', ''

    test 'one line', ->
      run 'DoCmd.Close', 'Close()'

    test 'bracketed', ->
      assert_js foo("""Me![Customer Orders].Requery
                       DoCmd.Close"""),
                Foo: """ns('Me').bang('Customer Orders').dot('Requery')();
                        ns('DoCmd').dot('Close')();\n"""

    # SKIPPED
    if 0 then test 'function().property', ->
      assert_js foo("""CurrentDb().Properties("StartupForm")"""),
                Foo: "ns('CurrentDb')().dot('Properties')('StartupForm');"

    test 'nested dot', ->
      assert_js foo('DoCmd.Nested.Close'),
                Foo: "ns('DoCmd').dot('Nested').dot('Close')();\n"

    test 'arguments', ->
      run 'DoCmd.OpenForm "Main Switchboard", 123',
          'OpenForm("Main Switchboard",123)'

    test 'arguments missing', ->
      run 'DoCmd.OpenReport "Sales by Category", 1, , 3',
          'OpenReport("Sales by Category",1,undefined,3)'

    test 'one argument', ->
      run 'DoCmd.OpenForm "Main Switchboard"',
          'OpenForm("Main Switchboard")'

    test 'one braced argument', ->
      run 'DoCmd.OpenForm ("Main Switchboard")',
          'OpenForm("Main Switchboard")'

    test 'two braced arguments', ->
      run 'DoCmd.MoveSize (1), (2)', 'MoveSize(1,2)'

    test 'dotted argument', ->
      assert_js foo('MsgBox A.B'), Foo: "ns('MsgBox')(ns('A').dot('B'));\n"

    test 'numbers', ->
      run 'DoCmd.Foo 1, 23, 456', 'Foo(1,23,456)'

    test 'comment', ->
        run "' hi there!", ''

    test 'double comment', ->
        run "' hi there!\n' bye!", ''

    test 'double comment indented', ->
        run "  ' hi there!\n  ' bye!", ''

    test 'end comment', ->
      run "DoCmd.Close ' this is comment", 'Close()'

    test 'complex indented', ->
      run "  ' Closes Startup form.\n" +
          "  ' Used in OnClick property...\n" +
          "   DoCmd.Close\n" +
          "   DoCmd.OpenForm (\"Main Switchboard\")",
          'Close()\nOpenForm("Main Switchboard")'

    test 'strange', ->
      run 'DoCmd.EndFunction', 'EndFunction()'

    test 'leading empty line', ->
      assert_js "\n#{foo 'DoCmd.Close'}", Foo: "ns('DoCmd').dot('Close')();\n"

    test 'empty module', ->
        runModule ''

    test 'declarations stub', ->
        test_foo_close before_func: """Option Compare Database
                                       Option Explicit
                                       Attribute VB_Exposed = False
                                       Dim path As String
                                       ' Decrarations end here"""

    test 'func_def arguments stub', ->
      assert_js """Sub Foo(A, B As Integer, ByVal C As Integer)
                     DoCmd.Close
                   End Sub""", Foo: "ns('DoCmd').dot('Close')();\n"

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
             DoCmd.WillNotRun""", 'Close()'

    test 'Label stub', ->
      run """On Error GoTo FooError
               DoCmd.Close
             FooExit:
               Exit Function
             FooError:
               DoCmd.HandleError
               Resume FooExit""", 'Close()'

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
      test_foo_close after: """If Err = 3270 Then
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
      assert_js """Function Foo()
                     DoCmd.Open
                   End Function
                   ' Second function
                   Private Sub Bar()
                     DoCmd.Close
                   End Sub""",
                Foo: "ns('DoCmd').dot('Open')();\n"
                Bar: "ns('DoCmd').dot('Close')();\n"
