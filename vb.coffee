parser = require "./parser"
escodegen = require "escodegen"

repr = (arg) -> require('util').format '%j', arg
pprint = (arg) -> console.log require('util').inspect arg, false, null

# parse VB expression or module, and return Parser API AST (for escodegen)
parse = (expr) ->
    tree = parser.parse expr

    # copy-pasted from sqld3/parse_sql.coffee
    Object.getPrototypeOf(tree).toString = (spaces) ->
        if not spaces then spaces = ""

        value = (if this.value? then "=> #{repr this.value}" else '')
        string = spaces + this.name +  " <" + repr(this.innerText()) + "> " + value
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
                            arguments: [ value ]
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
                when 'number'
                    literal parseInt(n.innerText(), 10)

                when 'module'
                    n.children[1].value
                when 'func_defs'
                    type: 'ObjectExpression'
                    properties: (value for {value} in n.children ? [])
                when 'func_def'
                    name = n.children[1].value
                    body = n.children[3].value
                    type: 'Property'
                    key:
                        type: 'Literal'
                        value: name
                    value:
                        type: 'FunctionExpression'
                        id: null
                        params: []
                        defaults: []
                        body:
                           type: 'BlockStatement'
                           body: body
                        rest: null
                        generator: false
                        expression: false
                    kind: 'init'
                when 'func_body'
                    for {value} in n.children when value?
                        value
                when 'statement'
                    n.children[0].value
                when 'exit_statement'
                    type: 'ReturnStatement'
                    argument: null
                when 'call_statement'
                    type: 'ExpressionStatement'
                    expression:
                        type: 'CallExpression'
                        callee: n.children[0].value
                        arguments: n.children[1]?.value ? []
                when 'callee'
                    for {value}, i in n.children by 2
                        expression =
                            type: 'CallExpression'
                            callee:
                                if op?
                                    type: 'MemberExpression'
                                    computed: no
                                    object: expression
                                    property:
                                        identifier {'.':'dot', '!':'bang'}[op]
                                else
                                    identifier 'scope'
                            arguments: [ literal value ]
                        op = n.children[i+1]?.value
                    expression
                when 'call_args'
                    for {value} in n.children[1..] by 2 then value
                when 'uname'
                    n.children[0].value
                when 'uname_itself'
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
        arguments: [literal func_name]
    arguments: [identifier('me'), identifier('us')].concat args

literal = (value) -> type: 'Literal', value: value
identifier = (name) -> type: 'Identifier', name: name

# generate JavaScript for a tree
generate = (tree) ->
    #console.log 'TREE:'
    #pprint tree
    body = "return #{escodegen.generate tree};"
    #console.log 'CODE =', "`" + body + "`"
    new Function 'me', 'us', 'fn', body

exports.compile = (expr) ->
    generate parse expr

exports.evaluate = (expr, me, us, fns) ->
    tree = parse expr
    if tree?
        js = generate tree
        fn_get = (name) ->
                 unless fns[name]?
                     throw new VBRuntimeError "VB function '#{name}' not found"
                 (args...) -> fns[name](args...)
        [me_get, us_get] = for obj in [me, us] then do (obj) ->
            (field) ->
                unless obj[field]?
                    throw new VBRuntimeError "VB field '#{field}' not found"
                obj[field]
        js me_get, us_get, fn_get
    else
        'Error parsing ' + expr

exports.loadmodule = (code, scope) ->
    tree = parse code
    if tree?
        #console.log 'TREE:'
        #pprint tree
        func = new Function 'scope', "return #{escodegen.generate tree};"
        #console.log 'CODE =', "`" + func + "`"
        func (name) ->
                unless scope[name]?
                    throw new VBRuntimeError "VB name '#{name}' not found"
                scope[name]
    else
        throw "Error parsing module '#{code[..100]}...'"

class VBRuntimeError extends Error
    constructor: (msg) ->
        @name = 'VBRuntimeError'
        @message = msg or @name
exports.VBRuntimeError = VBRuntimeError

# Usage: coffee vbjs.coffee "[foo]&[bar]"
#parse process.argv[2]
