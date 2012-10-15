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

run = (code, expected, use) ->
    log = ''
    code = "Function Foo()
            #{code}
            End Function"
    m = loadmodule code, DoCmd:
                            dot: (name) ->
                                (args...) ->
                                    spec = (repr a for a in args).join ','
                                    log += "#{name}(#{spec})\n"
    m.Foo()
    assert.strictEqual log, expected

suite 'Modules -', ->
    test 'empty', ->
        run '', ''

    test 'one line', ->
        run 'DoCmd.Close', 'Close()\n'

    test 'arguments', ->
        run 'DoCmd.OpenForm ("Main Switchboard")',
            'OpenForm("Main Switchboard")\n'

    test 'function', ->
        run """' Closes Startup form.
               ' Used in OnClick property of OK command button on Startup form.
                   DoCmd.Close
                   DoCmd.OpenForm ("Main Switchboard")
               End Function""", 'Close()\nOpenForm("Main Switchboard")\n'
