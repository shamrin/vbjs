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

logger = ->
    log = ''
    object:
        dot: (name) ->
            (args...) ->
                spec = (repr a for a in args).join ','
                log += "#{name}(#{spec})\n"
    log: ->
        log

suite 'Modules -', ->
    test 'empty', ->
        l = logger()
        m = loadmodule """Function CloseForm()
                          End Function""",
                       DoCmd: l
        m.CloseForm()
        assert.strictEqual l.log(), ''

    test 'basic', ->
        l = logger()
        m = loadmodule """Function CloseForm()
                              DoCmd.Close
                          End Function""",
                       DoCmd: l
        m.CloseForm()
        assert.strictEqual l.log(), """Close()\n"""

    test 'function', ->
        l = logger()
        m = loadmodule """Function CloseForm()
                          ' Closes Startup form.
                          ' Used in OnClick property of OK command button on Startup form.
                              DoCmd.Close
                              DoCmd.OpenForm ("Main Switchboard")
                          End Function""",
                       DoCmd: l
        m.CloseForm()
        assert.strictEqual l.log(), """Close()\nOpenForm("Main Switchboard")\n"""
