{runExpression, compileExpression, nsObject, VBRuntimeError} = require '../vb'
assert = require 'assert'

run = (expr, me={}, us={}, get_fns=null) ->
    namespace = _Us: {dotobj: us}, Me: {dotobj: me}
    if get_fns
      for k, v of get_fns nsObject {dotobj: namespace}
        namespace[k] = v
    runExpression expr, {dotobj: namespace}

eq = (expected, actual, msg) -> assert.strictEqual actual, expected, msg

# wrap primitive `value` with `VBObject`-like interface
obj = (value) -> get: -> value

nancy = FirstName: obj('Nancy'), LastName: obj('Davolio')

get_fns = (ns) ->
    Abs: (expr) -> obj Math.abs(expr)
    Sum: (expr) ->
        # TODO implement Sum in VBA, using CurrentDb.OpenRecordset or DBEngine
        field = {'[Field]': 'Field'}[expr]
        sum = 0
        for val in ns.dot('_Us').dot(field) then sum += val
        obj sum

suite 'Expressions -', ->
    setup ->
        # nothing here
    test 'basic', ->
        eq 'Nancy, Davolio', run '[FirstName]&", "&[LastName]', nancy
    test 'basic whitespace', ->
        eq 'Nancy Davolio', run '[FirstName] & " " & [LastName]', nancy
    test 'identifier with whitespace', ->
        eq 'Nancy', run '[First name]', 'First name': obj 'Nancy'
    test 'identifier operators', ->
        eq 'Total: 10', run '"Total: " & [Some Subform].[Form]![Subtotal]',
                        'Some Subform':
                          dotobj: Form: bang: (g) -> {Subtotal: obj 10}[g]
    test 'addition', ->
        eq 20, run '[Subtotal]+[Freight]',
                   Subtotal: obj 13
                   Freight: obj 7
    test 'functions', ->
        eq 30, run 'Abs([Field])', {Field: obj -30}, {}, get_fns
    test 'lazy functions', ->
        eq 40, run 'Sum([Field])', {}, {Field: [10, 10, 20]}, get_fns
    test 'arithmetic', ->
        eq 5, run '[A] / [B] + [C] * [D] - [E]',
              A: obj(9), B: obj(3), C: obj(2), D: obj(3), E: obj(4)
    test 'float', ->
        eq 2.0, run '[A] + 0.5', A: obj(1.5)
    test 'nested calls', ->
        eq 60, run 'Abs(Sum([Field]))', {}, {Field: [10, 20, -90]}, get_fns
    test 'unknown function error' , ->
        assert.throws (-> run 'Foo([F])', {F: 123}), VBRuntimeError
    test 'unknown me.field error', ->
        assert.throws (-> run '[Field]'), VBRuntimeError
    test 'unknown us.field error', ->
        assert.throws (-> run 'Sum([Field])', {}, {}, get_fns), VBRuntimeError
    test 'generated code', ->
        eq "ns = ns.dot; var me = ns('Me').dot; return me('Field').get();",
           compileExpression('[Field]').toString()
    test 'case insensitive function name', ->
        eq 30, run 'aBs([Field])', {Field: obj -30}, {}, get_fns
