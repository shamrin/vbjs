# Other tests to add (from nwind.mdb):
#    'Sum([ExtendedPrice])'
#    '[Quarterly Orders Subform]![Total]'
#    '"Grand Total for " & [Forms]![Quarterly Orders]![Quarterly Orders Subform].[Form]![Year]'
#    'NZ(Sum([Qtr 1]))'
#    '[TotalQ1]+[TotalQ2]+[TotalQ3]+[TotalQ4]'

{evaluate} = require '../vbjs'
{strictEqual} = require 'assert'

run = (expr, Me) -> evaluate expr, Me
eq = (expected, actual, message) -> strictEqual actual, expected, message

suite 'Expressions -', ->
    setup ->
        # nothing here
    test 'basic', ->
        eq 'Nancy, Davolio', run '[FirstName]&", "&[LastName]',
                             FirstName: 'Nancy', LastName: 'Davolio'
    test 'basic whitespace', ->
        eq 'Nancy Davolio', run '[FirstName] & " " & [LastName]',
                            FirstName: 'Nancy', LastName: 'Davolio'
    test 'identifier with whitespace', ->
        eq 'Nancy', run '[First name]', 'First name': 'Nancy'
    test 'identifier operators', ->
        eq 123, run '[Some Subform].[Form]![Subtotal]',
                    'Some Subform':
                        Form: __default: (v) -> {Subtotal: 123}[v]
    test 'addition', ->
        eq 666, run '[Subtotal]+[Freight]',
                    Subtotal: 123, Freight: 543
