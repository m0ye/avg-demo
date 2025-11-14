extends Object
# 对话数据类
class_name DialogueNode

## 对话节点类型
enum NodeType {
	dialogue, ## 对话类型
	choice, ## 选项类型
	command, ## 命令类型
	narration ## 旁白类型
}

## 选项数据类
class Choice:
	var choice_name: String
	var choice_text: String
	var choice_next_node: String


## 对话节点名称
var node: String

## 对话节点类型 
var node_type: NodeType

## 对话角色，当节点类型为对话类型时起效
var character: String

## 对话或旁白文本
var text: String

## 下一节点，对话和旁白类型起效
var next_node: String

## 选项对象，选项类型起效
var choice: Array[Choice]

## 指令，指令类型起效
var command: String