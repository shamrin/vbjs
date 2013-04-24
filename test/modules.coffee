{isArray, isRegExp} = require 'underscore'
{runModule, compileModule, runJS, VBObject, VBRuntimeError} = require '../vb'
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
        if isArray o
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
  module = runModule code,
                     new VBObject
                       attrs: DoCmd: new VBObject
                         dot: (name) ->
                                (args...) ->
                                  spec = (repr a for a in args).join ','
                                  log.push "#{name}(#{spec})"
  module.Foo()
  assert.strictEqual log.join('\n'), expected
  module

assert_js = (vba, expected_obj) ->
  actual = runJS compileModule vba
  for fn, expected of expected_obj
    if isRegExp expected
      assert.ok actual[fn].toString().match(expected),
                "#{expected} doesn't match '#{actual[fn]}'"
    else
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

  assert_js code, Foo: "ns.get('DoCmd').dot('Close')();\n"

foo = (code) -> "Function Foo()\n  #{code}\nEnd Function"

suite 'Modules -', ->
  test 'empty', ->
    run '', ''

  test 'one line', ->
    run 'DoCmd.Close', 'Close()'

  test 'bracketed', ->
    assert_js foo("""Me![Customer Orders].Requery
                     DoCmd.Close"""),
              Foo: """ns.get('Me').bang('Customer Orders').dot('Requery')();
                      ns.get('DoCmd').dot('Close')();\n"""

  # SKIPPED
  if 0 then test 'function().property', ->
    assert_js foo("""CurrentDb().Properties("StartupForm")"""),
              Foo: "ns.get('CurrentDb')().dot('Properties')('StartupForm');"

  test 'nested dot', ->
    assert_js foo('DoCmd.Nested.Close'),
              Foo: "ns.get('DoCmd').dot('Nested').dot('Close')();\n"

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
    assert_js foo('MsgBox A.B'), Foo: "ns.get('MsgBox')(ns.get('A').dot('B').get());\n"

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
    assert_js "\n#{foo 'DoCmd.Close'}", Foo: "ns.get('DoCmd').dot('Close')();\n"

  test 'empty module', ->
    runModule ''

  test 'declarations stub', ->
    test_foo_close before_func: """Option Compare Database
                                     Option Explicit
                                     Attribute VB_Exposed = False
                                     Dim path As String
                                     ' Decrarations end here"""

  test 'func_def arguments', ->
    assert_js """Sub Foo(S, M As Integer, ByVal N As Integer)
                   DoCmd.Bar S, M
                 End Sub""",
              Foo: /^function \(S, M, N\) \{\s*ns\.get\('DoCmd'\)\.dot\('Bar'\)\(S, M\);\s*\}$/

  test 'function As stub', ->
    test_foo_close after_spec: 'As Boolean'

  test 'function Error Resume stub', ->
    test_foo_close before: 'On Error Resume Next'

  test 'function Error GoTo stub', ->
    test_foo_close before: 'On Error GoTo LabelName'

  test 'Resume stub', ->
    test_foo_close before: 'On Error GoTo LabelName',
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
                              Dim Num"""

  test 'assign', ->
    assert_js foo("""foo = "Foo"
                     bar = False"""),
              Foo: """ns.get('foo').let('Foo');
                      ns.get('bar').let(false);\n"""

  test 'complex assign', ->
    assert_js foo('CurrentDb().Properties("StartupForm") = "Form1"'),
              Foo: "ns.get('CurrentDb')().dot('Properties')('StartupForm').let('Form1');\n"

  test 'assign real-life', ->
    assert_js foo("""If [type] = 1 Then
                       n1.Visible = True
                     Else
                       n1.Visible = False
                     End If"""),
              Foo: """if (ns.get('type').get() === 1) {
                      ns.get('n1').dot('Visible').let(true);
                      } else {
                      ns.get('n1').dot('Visible').let(false);
                      }\n"""

  test 'simple If', ->
    assert_js foo("""If Bar Then
                       DoCmd.Bla
                     End If"""),
              Foo: """if (ns.get('Bar').get()) {
                      ns.get('DoCmd').dot('Bla')();
                      }\n"""

  test 'Like', -> # TODO move to test/expressions.coffee
    assert_js foo("""If Foo Like "[A-Z]" Then
                       DoCmd.Bla
                     End If"""),
              Foo: """if (/[A-Z]/.test(ns.get('Foo').get())) {
                      ns.get('DoCmd').dot('Bla')();
                      }\n"""

  test 'If Not EndIf', ->
    assert_js foo("""If Not IsNull(Me!Photo) Then
                       hideImageFrame
                     EndIf"""),
              Foo: """if (!ns.get('IsNull')(ns.get('Me').bang('Photo').get()).get()) {
                      ns.get('hideImageFrame')();
                      }\n"""

  test 'If ElseIf Else', ->
    assert_js foo("""If IsIt() Then
                       DoCmd.DoA
                     ElseIf IsLoaded("Product List") Then
                       DoCmd.OpenForm strDocName
                     Else
                       DoCmd.DoA
                       DoCmd.DoB
                     End If"""),
              Foo: """if (ns.get('IsIt')().get()) {
                      ns.get('DoCmd').dot('DoA')();
                      } else if (ns.get('IsLoaded')('Product List').get()) {
                      ns.get('DoCmd').dot('OpenForm')(ns.get('strDocName').get());
                      } else {
                      ns.get('DoCmd').dot('DoA')();
                      ns.get('DoCmd').dot('DoB')();
                      }\n"""

  test 'If single line', ->
    assert_js foo("If IsItReplica() Then DoCmd.TellAboutIt"),
              Foo: """if (ns.get('IsItReplica')().get())
                      ns.get('DoCmd').dot('TellAboutIt')();\n"""

  test 'Condition in braces', ->
    assert_js foo("""If (IsItReplica()) Then
                       DoCmd.TellAboutIt
                     End If"""),
              Foo: """if (ns.get('IsItReplica')().get()) {
                      ns.get('DoCmd').dot('TellAboutIt')();
                      }\n"""

  test 'If Or _', ->
    assert_js foo("""If (CurrentDb().Properties("Form1") = "Startup" Or _
                     CurrentDb().Properties("Form1") = "Form.Startup") Then
                       Forms!Startup!HideStartupForm = False
                     End If"""),
              Foo: """if (ns.get('CurrentDb')().dot('Properties')('Form1').get() === 'Startup' || ns.get('CurrentDb')().dot('Properties')('Form1').get() === 'Form.Startup') {
                       ns.get('Forms').bang('Startup').bang('HideStartupForm').let(false);
                       }\n"""

  test 'If Or And', ->
    assert_js foo("""If (Aa = "") Or (0 < 1) _
                         Or (1 <> 2) And (D >= 5) Then
                       Bar
                     End If"""),
              Foo: """if (ns.get('Aa').get() === '' || 0 < 1 || 1 !== 2 && ns.get('D').get() >= 5) {
                      ns.get('Bar')();
                      }\n"""

  test 'If Or CrLf', ->
    assert_js foo("If A _\r\nOr B Then\r\nC = 0\r\nEnd If"),
              Foo: """if (ns.get('A').get() || ns.get('B').get()) {
                      ns.get('C').let(0);
                      }\n"""

  test 'Select Case stub', ->
    test_foo_close before: """Select Case Me!ReportToPrint
                                Case 1, 3 To 5
                                  DoCmd.OpenReport "Sales report", PrintMode
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
              Foo: "ns.get('DoCmd').dot('Open')();\n"
              Bar: "ns.get('DoCmd').dot('Close')();\n"
