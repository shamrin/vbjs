parser = require "./vb.js"
escodegen = require "escodegen"

exports.evaluate = (expr, Me, Us, functions) ->
    tree = parse expr
    if tree?
        js = compile tree
        js Me, Us, functions
    else
        'Error parsing ' + expr

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
                    type: 'Literal'
                    value: n.children[1].value
                when 'literal_text'
                    n.innerText()
                when 'identifier_expr'
                    result =
                        type: 'Identifier'
                        name: 'Me'
                    for {value}, i in n.children by 2
                        result = switch op ? '.'
                            when '.' # A.B => A[B]
                                type: 'MemberExpression'
                                computed: yes
                                object: result
                                property: value
                            when '!'
                                # bang op: A!B => A.__default(B), links:
                                # * http://stackoverflow.com/q/4804947
                                # * http://stackoverflow.com/q/2923957
                                # * http://www.cpearson.com/excel/DefaultMember.aspx
                                type: 'CallExpression'
                                callee: 
                                    type: 'MemberExpression'
                                    computed: no
                                    object: result
                                    property:
                                        type: 'Identifier'
                                        name: '__default'
                                'arguments': [ value ]
                        op = n.children[i+1]?.value
                    result
                when 'identifier_op'
                    n.innerText()
                when 'identifier'
                    type: 'Literal'
                    value: n.children[1].value
                when 'name_in_brackets', 'lazy_name'
                    n.innerText()
                when 'concat_expr' 
                    result = if n.children[1]? # force string
                                 type: 'Literal', value: ''
                    for {value}, i in n.children by 2
                        result = if result? then plus result, value else value
                    result
                when 'add_expr'
                    for {value}, i in n.children by 2
                        result = if result? then plus result, value else value
                    result

                when 'lazy_call_expr'
                    [{value: func_name}, l, params..., r] = n.children
                    type: 'CallExpression'
                    callee:
                        type: 'MemberExpression'
                        computed: 'true'
                        object:
                            type: 'Identifier'
                            name: 'functions'
                        property:
                            type: 'Literal'
                            value: func_name
                    'arguments': [{type: 'Identifier', name: 'Me'},
                                  {type: 'Identifier', name: 'Us'}]\
                                 .concat(for {value} in params by 2
                                             type: 'Literal'
                                             value: value)
                when 'lazy_value'
                    n.innerText()

            #if n.name is 'start' then console.log n.toString()

    #pprint tree
    tree.value

plus = (left, right) ->
    type: 'BinaryExpression'
    operator: '+'
    left: left
    right: right

# compile to JavaScript
compile = (tree) ->
    #console.log 'TREE:'
    #pprint tree
    code = escodegen.generate tree
    #console.log 'CODE', code
    new Function 'Me', 'Us', 'functions', "return #{code};"

# Usage: coffee vbjs.coffee "[foo]&[bar]"
#parse process.argv[2]
