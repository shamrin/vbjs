parser = require "./parser"
escodegen = require "escodegen"

repr = (arg) -> require('util').format '%j', arg
pprint = (arg) -> console.log require('util').inspect arg, false, null

member = (op) -> {'.': 'dot', '!': 'bang'}[op]

# Parse VB expression or module, and return Parser API AST [1] for escodegen.
#
# Escodegen [2] will translate this AST to JavaScript. To figure out how to
# construct a certain JavaScript fragment, use parser demo [3].
#
# [1]: https://developer.mozilla.org/en/SpiderMonkey/Parser_API
# [2]: https://github.com/Constellation/escodegen
# [3]: http://esprima.org/demo/parse.html
parse = (expr) ->
    tree = parser.parse expr

    # first copy-pasted from sqld3/parse_sql.coffee
    Object.getPrototypeOf(tree).toString = (spaces = '') ->
        result = try
                     "=> #{escodegen.generate @value}"
                 catch error
                     if @value? then "=> AST: #{repr @value}" else ''

        string = spaces + "#{@name} <#{repr @innerText()}> " + result
        for child in @children when typeof child isnt 'string'
            string += "\n" + child.toString(spaces + ' ')

        return string

    tree.traverse
        traversesTextNodes: false
        exitedNode: (n) ->
            n.value = switch n.name
                when '#document', 'source' # language.js nodes
                    n.children?[0]?.value
                when 'start'
                    n.children[1].value
                when 'expression', 'value', 'identifier', 'identifier_expr_part'
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
                    result = n.children[0].value
                    if result.type is 'Literal'
                        result =
                            type: 'CallExpression'
                            callee: identifier 'me'
                            arguments: [ result ]
                    for {value: arg}, i in n.children by 2 when i > 0
                        result =
                            type: 'CallExpression'
                            callee:
                                type: 'MemberExpression'
                                computed: no
                                object: result
                                property:
                                    identifier member n.children[i-1].value
                            arguments: [ arg ]
                    result
                when 'identifier_op'
                    n.innerText()
                when 'bracketed_identifier'
                    literal n.children[1].value
                when 'name', 'name_in_brackets', 'lazy_name'
                    n.innerText()
                when 'or_expr', 'cmp_expr'
                    n.children[0].value # FIXME it's just a stub now
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
                    n.children[0].value
                when 'plain_call_expr'
                    [{value: fn}, l, params..., r] = n.children
                    call(fn, for {value} in params by 2 then value)
                when 'lazy_call_expr'
                    [{value: fn}, l, params..., r] = n.children
                    call(fn, for {value} in params by 2 then literal value)
                when 'lazy_value'
                    n.innerText()
                when 'number'
                    literal parseInt(n.innerText(), 10)

                when 'module'
                    n.children[2].value
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
                when 'statements'
                    for {value} in n.children when value?
                        value
                when 'single_line_statement'
                    n.children[0].value
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
                    expression =
                        type: 'CallExpression'
                        callee: identifier 'scope'
                        arguments: [ literal n.children[0].value ]
                    for {value: arg}, i in n.children by 2 when i > 0
                        expression =
                            type: 'CallExpression'
                            callee:
                                type: 'MemberExpression'
                                computed: no
                                object: expression
                                property:
                                    identifier member n.children[i-1].value
                            arguments: [ literal arg ]
                    expression
                when 'call_spec'
                    n.children[1]?.value ? n.children[0].value
                when 'call_args'
                    for {value} in n.children by 2 then value
                when 'uname'
                    n.children[0].value
                when 'uname_itself'
                    n.innerText()

            #if n.name is 'start' then console.log n.toString()

    if not tree.value? and process?.env?.TESTING?
        require('./test/pegjs_parser').check '<string>', expr

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
    generate parse 'E' + expr

exports.evaluate = (expr, me, us, fns) ->
    tree = parse 'E' + expr
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

# load VBA module and return object with VB functions, compiled to JavaScript
exports.loadmodule = (code, scope) ->
    tree = parse 'M' + code
    if tree?
        #console.log 'TREE:'
        #pprint tree
        body = "return #{escodegen.generate tree};"
        try
            func = new Function 'scope', body
        catch error
            console.log "#{error} in `#{body}`"
            throw error
        #console.log 'CODE =', "`" + func + "`"
        func (name) ->
                unless scope[name]?
                    throw new VBRuntimeError "VB name '#{name}' not found"
                scope[name]
    else
        throw "Error parsing module '#{code[..150]}...'"

class VBRuntimeError extends Error
    constructor: (msg) ->
        @name = 'VBRuntimeError'
        @message = msg or @name
exports.VBRuntimeError = VBRuntimeError

# Usage: coffee vbjs.coffee "[foo]&[bar]"
#parse process.argv[2]
