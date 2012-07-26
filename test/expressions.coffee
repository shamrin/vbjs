# Other tests to add (from nwind.mdb):
#    '[Orders Subform].[Form]![OrderSubtotal]'
#    '[Subtotal]+[Freight]'
#    'Sum([ExtendedPrice])'
#    '[Quarterly Orders Subform]![Total]'
#    '"Grand Total for " & [Forms]![Quarterly Orders]![Quarterly Orders Subform].[Form]![Year]'
#    'NZ(Sum([Qtr 1]))'
#    '[TotalQ1]+[TotalQ2]+[TotalQ3]+[TotalQ4]'

{evaluate} = require '../vbjs'
{strictEqual} = require 'assert'

run = (expr, Me) -> evaluate expr, Me
eq = (expected, actual, message) -> strictEqual actual, expected, message

suite 'Expressions', ->
    setup ->
        # nothing here
    test 'basic', ->
        eq 'Nancy, Davolio', run '[FirstName]&", "&[LastName]',
                             FirstName: 'Nancy', LastName: 'Davolio'
    test 'basic whitespace', ->
        eq 'Nancy Davolio', run '[FirstName] & " " & [LastName]',
                            FirstName: 'Nancy', LastName: 'Davolio'
    test 'identifier operators', ->
        eq 123, run '[SomeSubform].[Form]![Subtotal]',
                    SomeSubform:
                        Form:
                            __default: (v) -> {Subtotal: 123}[v]

