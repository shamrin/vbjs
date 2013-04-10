{runExpression, compileExpression, VBObject, VBRuntimeError} = require '../vb'
assert = require 'assert'

run = (expr, me={}, us={}, get_fns=null) ->
    ns =
      _Us: new VBObject {attrs: us, type: '_UsType'}
      Me: new VBObject {attrs: me, type: 'MeType'}
    if get_fns
      for k, v of get_fns new VBObject {attrs: ns}
        ns[k] = v
    runExpression expr, new VBObject {attrs: ns}

eq = (expected, actual, msg) -> assert.strictEqual actual, expected, msg

# wrap primitive `value` with `VBObject`-like interface
obj = (value) -> get: -> value

nancy = FirstName: 'Nancy', LastName: 'Davolio'

get_fns = (ns) ->
    Abs: (expr) -> obj Math.abs(expr)
    Sum: (expr) ->
        # TODO implement Sum in VBA, using CurrentDb.OpenRecordset or DBEngine
        field = {'[Field]': 'Field'}[expr]
        sum = 0
        for val in ns.dot('_Us').get(field) then sum += val
        obj sum

suite 'Expressions -', ->
    setup ->
        # nothing here
    test 'basic', ->
        eq 'Nancy, Davolio', run '[FirstName]&", "&[LastName]', nancy
    test 'basic whitespace', ->
        eq 'Nancy Davolio', run '[FirstName] & " " & [LastName]', nancy
    test 'identifier with whitespace', ->
        eq 'Nancy', run '[First name]', 'First name': 'Nancy'
    test 'identifier operators', ->
        eq 'Total: 10', run '"Total: " & [Some Subform].[Form]![Subtotal]',
                        'Some Subform': new VBObject
                            attrs: Form: new VBObject
                                bang: (g) -> {Subtotal: obj 10}[g]
    test 'addition', ->
        eq 20, run '[Subtotal]+[Freight]', Subtotal: 13, Freight: 7
    test 'function', ->
        eq 30, run 'Abs([Field])', {Field: -30}, {}, get_fns
    test 'insensitive function', ->
        eq 30, run 'aBs([Field])', {Field: -30}, {}, get_fns
    test 'lazy functions', ->
        eq 40, run 'Sum([Field])', {}, {Field: [10, 10, 20]}, get_fns
    test 'arithmetic', ->
        eq 5, run '[A] / [B] + [C] * [D] - [E]', A: 9, B: 3, C: 2, D: 3, E: 4
    test 'float', ->
        eq 2.0, run '[A] + 0.5', A: 1.5
    test 'nested calls', ->
        eq 60, run 'Abs(Sum([Field]))', {}, {Field: [10, 20, -90]}, get_fns
    test 'unknown function error' , ->
        assert.throws (-> run 'Foo([F])', {F: 123}), VBRuntimeError
    test 'unknown me.field error', ->
        assert.throws (-> run '[Field]'), VBRuntimeError
    test 'unknown us.field error', ->
        assert.throws (-> run 'Sum([Field])', {}, {}, get_fns), VBRuntimeError
    test 'generated code', ->
        eq "var me = ns.get('Me'); return me.dot('Field').get();",
           compileExpression('[Field]').toString()
    test 'generated function code', ->
        eq "var me = ns.get('Me'); return ns.get('Abs')(me.dot('Field').get()).get();",
           compileExpression('Abs([Field])').toString()
