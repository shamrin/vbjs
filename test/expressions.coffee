# Other tests to add (from nwind.mdb):
#    'NZ(Sum([Qtr 1]))'

{evaluate, compile, VBRuntimeError} = require '../vbjs'
assert = require 'assert'

run = (expr, me={}, us={}, fns={}) -> evaluate expr, me, us, fns
eq = (expected, actual, msg) -> assert.strictEqual actual, expected, msg

nancy = FirstName: 'Nancy', LastName: 'Davolio'
fns =
    Abs: (me, us, expr) -> Math.abs(expr)
    Sum: (me, us, expr) ->
        field = {'[Field]': 'Field'}[expr]
        sum = 0
        for val in us(field) then sum += val
        sum

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
                        'Some Subform':
                            dot: (f) -> {Form:
                                            bang: (g) -> {Subtotal: 10}[g]}[f]
    test 'addition', ->
        eq 20, run '[Subtotal]+[Freight]', Subtotal: 13, Freight: 7
    test 'functions', ->
        eq 30, run 'Abs([Field])', {Field: -30}, {}, fns
    test 'lazy functions', ->
        eq 40, run 'Sum([Field])', {}, {Field: [10, 10, 20]}, fns
    test 'long addition', ->
        eq 50, run '[X] + [Y] + [Z]', X: 10, Y: 20, Z: 20
    test 'nested calls', ->
        eq 60, run 'Abs(Sum([Field]))', {}, {Field: [10, 20, -90]}, fns
    test 'unknown function error' , ->
        assert.throws (-> run 'Foo([F])', {F: 123}), VBRuntimeError
    test 'unknown me.field error', ->
        assert.throws (-> run '[Field]'), VBRuntimeError
    test 'unknown us.field error', ->
        assert.throws (-> run 'Sum([Field])', {}, {}, fns), VBRuntimeError
    test 'generated code', ->
        eq "function anonymous(me,us,fn) {\nreturn me('Field');\n}",
           compile('[Field]').toString()
