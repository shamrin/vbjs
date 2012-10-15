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

assert_js = (module, expected) ->
    match = module.Foo.toString().match /^function \(\) \{\s*(.*)\s*\}$/m
    assert.strictEqual match[1], expected

test_foo_close = ({before, after, after_spec, before_func}) ->
    if before? then before += '\n' else before = ''
    if after? then after = '\n' + after else after = ''
    if after_spec? then after_spec += '\n' else after_spec = ''
    if before_func? then before_func += '\n' else before_func = ''
    m = runmod """#{before_func}Function Foo() #{after_spec}
                    #{before}DoCmd.Close#{after}
                  End Function"""
    assert_js m, "scope('DoCmd').dot('Close')();"

suite 'Modules -', ->
    test 'empty', ->
        run '', ''

    test 'one line', ->
        run 'DoCmd.Close', 'Close()\n'

    test 'nested dot', ->
        assert_js run('DoCmd.Nested.Close'),
                  "scope('DoCmd').dot('Nested').dot('Close')();"

    test 'arguments', ->
        run 'DoCmd.OpenForm ("Main Switchboard")',
            'OpenForm("Main Switchboard")\n'

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
        assert_js m, "scope('DoCmd').dot('Close')();"

    test 'empty module', ->
        runmod ''

    test 'option stub', ->
        test_foo_close before_func: """Option Compare Database
                                       Option Explicit"""

    test 'function As stub', ->
        test_foo_close after_spec: 'As Boolean'

    test 'function As Error stub', ->
        test_foo_close
            after_spec: 'As Boolean'
            before: 'On Error GoTo LabelName'

    test 'function Error stub', ->
        test_foo_close before: 'On Error GoTo LabelName'

    test 'Resume stub', ->
        test_foo_close
            before: 'On Error GoTo LabelName'
            after: 'Resume LabelName'

    test 'Exit stub', ->
        test_foo_close after: 'Exit Function'

    test 'Exit Function', ->
        test_foo_close after: """Exit Function
                                 DoCmd.WillNotRun"""

    test 'Label stub', ->
        test_foo_close
            before: 'On Error GoTo FooError'
            after: """FooExit:
                          Exit Function
                      FooError:
                          DoCmd.HandleError
                          Resume FooExit"""

    test 'Const Set stub', ->
        test_foo_close before: """Const someConst = 42
                                  Set someVal = 43"""

    test 'Dim stub', ->
        test_foo_close before: 'Dim returnValue As Boolean'

    test 'assign stub', ->
        test_foo_close before: """Dim returnValue As Boolean
                                  returnValue = False"""

    test 'If stub', ->
        test_foo_close after: """FooErr:
                                 If Err = 3270 Then
                                     DoCmd.Bla
                                     Resume FooExit
                                 End If"""

    test 'If Or stub'
