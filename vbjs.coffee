parser = require "./vb.js"
escodegen = require "escodegen"

repr = (arg) -> require('util').format '%j', arg
pprint = (arg) -> console.log require('util').inspect arg, false, null

parse = (expr) ->
    tree = parser.parse expr

    # copy-pasted from sqld3/parse_sql.coffee
    Object.getPrototypeOf(tree).toString = (spaces) ->
        if not spaces then spaces = ""

        value = (if this.value? then "=> #{repr this.value}" else '')
        string = spaces + this.name +  " <" + this.innerText() + "> " + value
        children = this.children
        index = 0

        for child in children
            if typeof child == "string"
                #string += "\n" + spaces + ' ' + child
            else
                string += "\n" + child.toString(spaces + ' ')

        return string

    tree.traverse
        traversesTextNodes: false
        exitedNode: (n) ->
            n.value = switch n.name
                when '#document', 'source' # language.js nodes
                    n.children?[0]?.value
                when 'start', 'expression', 'value'
                    n.children[0].value
                when 'literal'
                    literal n.children[1].value
                when 'literal_text'
                    n.innerText()
                when 'identifier_expr'
                    # [A].[B]![C] => me('A').dot('B').bang('C')
                    # About bang ! operator semantics:
                    #   * http://stackoverflow.com/q/4804947
                    #   * http://stackoverflow.com/q/2923957
                    #   * http://www.cpearson.com/excel/DefaultMember.aspx
                    for {value}, i in n.children by 2
                        result =
                            type: 'CallExpression'
                            callee:
                                if op?
                                    type: 'MemberExpression'
                                    computed: no
                                    object: result
                                    property: 
                                        identifier {'.':'dot', '!':'bang'}[op]
                                else
                                    identifier 'me'
                            'arguments': [ value ]
                        op = n.children[i+1]?.value
                    result
                when 'identifier_op'
                    n.innerText()
                when 'identifier'
                    literal n.children[1].value
                when 'name', 'name_in_brackets', 'lazy_name'
                    n.innerText()
                when 'concat_expr' 
                    result = if n.children[1]? then literal '' # force string
                    for {value}, i in n.children by 2
                        result = if result? then plus result, value else value
                    result
                when 'add_expr'
                    for {value}, i in n.children by 2
                        result = if result? then plus result, value else value
                    result
                when 'call_expr'
                    if n.children.length > 1
                        [{value: fn}, l, params..., r] = n.children
                        call(fn, for {value} in params by 2 then value)
                    else
                        n.children[0].value
                when 'lazy_call_expr'
                    [{value: fn}, l, params..., r] = n.children
                    call(fn, for {value} in params by 2 then literal value)
                when 'lazy_value'
                    n.innerText()

            #if n.name is 'start' then console.log n.toString()

    #pprint tree
    tree.value

# `left` + `right`
plus = (left, right) ->
    type: 'BinaryExpression'
    operator: '+'
    left: left
    right: right

# fn("`func_name`")(me, us, `args`...)
call = (func_name, args) ->
    type: 'CallExpression'
    callee:
        type: 'CallExpression'
        callee: identifier 'fn'
        'arguments': [literal func_name]
    'arguments': [identifier('me'), identifier('us')].concat args

literal = (value) -> type: 'Literal', value: value
identifier = (name) -> type: 'Identifier', name: name

# compile to JavaScript
compile = (tree) ->
    #console.log 'TREE:'
    #pprint tree
    body = "return #{escodegen.generate tree};"
    #console.log 'CODE =', "`" + body + "`"
    new Function 'me', 'us', 'fn', body

exports.evaluate = (expr, me, us, fns) ->
    tree = parse expr
    if tree?
        js = compile tree
        fn_get = (name) ->
                 unless fns[name]?
                     throw new VBRuntimeError "VB function '#{name}' not found"
                 (args...) -> fns[name](args...)
        [me_get, us_get] = for obj in [me, us]
            do (obj) ->
                (field) ->
                    unless obj[field]?
                        throw new VBRuntimeError "VB field '#{field}' not found"
                    obj[field]
        js me_get, us_get, fn_get
    else
        'Error parsing ' + expr

class VBRuntimeError extends Error
    constructor: (msg) ->
        @name = 'VBRuntimeError'
        @message = msg or @name
exports.VBRuntimeError = VBRuntimeError

# Usage: coffee vbjs.coffee "[foo]&[bar]"
#parse process.argv[2]
