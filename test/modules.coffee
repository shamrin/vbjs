{loadmodule, VBRuntimeError} = require '../vb'
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
    module = loadmodule code, DoCmd:
                            dot: (name) ->
                                (args...) ->
                                    spec = (repr a for a in args).join ','
                                    log += "#{name}(#{spec})\n"
    if expected?
        module.Foo()
        assert.strictEqual log, expected
    module

runmod = (code) ->
    loadmodule code

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

    m = runmod """#{before_func}Function Foo() #{after_spec}
                    #{before}DoCmd.Close
                    #{after}End Function"""
    assert_js m, Foo: "scope('DoCmd').dot('Close')();\n"

suite 'Modules -', ->
    test 'empty', ->
        run '', ''

    test 'one line', ->
        run 'DoCmd.Close', 'Close()\n'

    # SKIPPED
    if 0 then test 'function().property', ->
        m = runmod """CurrentDb().Properties("StartupForm")"""
        assert_js m, Foo: "scope('CurrentDb')().dot('Properties')('StartupForm');"

    test 'nested dot', ->
        assert_js run('DoCmd.Nested.Close'),
                  Foo: "scope('DoCmd').dot('Nested').dot('Close')();\n"

    test 'arguments', ->
        run 'DoCmd.OpenForm ("Main Switchboard")',
            'OpenForm("Main Switchboard")\n'

    test 'arguments without braces', ->
        run 'DoCmd.OpenForm "Main Switchboard"',
            'OpenForm("Main Switchboard")\n'

    test 'numbers', ->
        run 'DoCmd.Foo (1, 23, 456)',
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
        m = runmod "\nFunction Foo()\nDoCmd.Close\nEnd Function"
        assert_js m, Foo: "scope('DoCmd').dot('Close')();\n"

    test 'empty module', ->
        runmod ''

    test 'option stub', ->
        test_foo_close before_func: """Option Compare Database
                                       Option Explicit
                                       ' Options end here"""

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

    test 'Const Set stub', ->
        test_foo_close before: """Const someConst = 42
                                  Set someVal = 43"""

    test 'Dim stub', ->
        test_foo_close before: """Dim returnValue As Boolean
                                  Static db As DAO.Database"""

    test 'assign stub', ->
        test_foo_close before: """foo = "Foo"
                                  bar = False"""

    test 'If stub', ->
        test_foo_close after: """FooErr:
                                 If Err = 3270 Then
                                     DoCmd.Bla
                                     Resume FooExit
                                 End If"""

    test 'If Else stub', ->
        test_foo_close before: """If IsItReplica() Then
                                      DoCmd.TellAboutIt
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

    test 'several functions', ->
        m = runmod """Function Foo()
                          DoCmd.Open
                      End Function
                      ' Second function
                      Private Sub Bar()
                          DoCmd.Close
                      End Sub"""
        assert_js m,
            Foo: "scope('DoCmd').dot('Open')();\n"
            Bar: "scope('DoCmd').dot('Close')();\n"

    test 'attribute stub', ->
        test_foo_close before_func: """Attribute VB_Exposed = False"""
