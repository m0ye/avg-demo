extends Node
# 配置管理器类，设定为全局单类、自动加载


## 对话资源路径
const DIALOGUE_DATA_PATH = "res://Configuration/dialogue.json"

## 选项资源路径
const CHOICES_DATA_PATH = "res://Configuration/choices.json"

## 对话数据
var dialogue_data: Array[DialogueNode]


# 构造函数加载资源数据
func _init() -> void:
	load_dialogue_file()

## 加载对话资源
func load_dialogue_file() -> bool:
	# 打开对话和选项文件
	var dialogue_file = FileAccess.open(DIALOGUE_DATA_PATH, FileAccess.READ)
	var choices_file = FileAccess.open(CHOICES_DATA_PATH, FileAccess.READ)
	if !dialogue_file or !choices_file:
		if !dialogue_file:
			push_error("无法打开对话数据文件：" + DIALOGUE_DATA_PATH)
		if !choices_file:
			push_error("无法打开选项数据文件：" + CHOICES_DATA_PATH)
		return false

	# 解析对话文本
	var dialogue_str_parse = parse_file(dialogue_file)
	if !dialogue_str_parse:
		push_error("对话数据JSON解析错误！")
		return false
	if not (dialogue_str_parse is Array):
		push_error("对话数据JSON根节点必须是数组！")
		return false

	# 解析选项文本
	var choices_str_parse = parse_file(choices_file)
	if !choices_str_parse:
		push_error("选项数据JSON解析错误！")
		return false
	if not (choices_str_parse is Array):
		push_error("选项数据JSON根节点必须是数组！")
		return false
	
	# 清空旧数据
	dialogue_data.clear()

	# 存入数据
	store_dialogue_data(dialogue_str_parse, choices_str_parse)

	return true

## 解析文件
func parse_file(file: FileAccess) -> Variant:
	# 读取文件数据转为string
	var file_str = file.get_as_text()
	file.close()

	# 解析string
	var str_parse = JSON.parse_string(file_str)
	return str_parse

## 处理数据存入dialogue_data
func store_dialogue_data(dialogue_parse: Variant, choices_parse: Variant):
	# 遍历每条对话数据
	for data in dialogue_parse:
		# 排除非字典的数据
		if not (data is Dictionary):
			continue

		var dialogue = DialogueNode.new()
		dialogue.node = data.get("node")
		dialogue.node_type = DialogueNode.NodeType[data.get("nodeType")]
		dialogue.character = data.get("character")
		dialogue.text = data.get("text")
		dialogue.next_node = data.get("nextNode")
		dialogue.command = data.get("command")

		# 处理选项类型数据
		if dialogue.node_type == DialogueNode.NodeType.choice:
			var choice_name = data.get("choice")

			for choice_data in choices_parse:
				if not (choice_data is Dictionary):
					continue
				if choice_name == choice_data.get("choiceName"):
					var choice = DialogueNode.Choice.new()
					choice.choice_name = choice_data.get("choiceName")
					choice.choice_text = choice_data.get("choiceText")
					choice.choice_next_node = choice_data.get("choiceNextNode")

					dialogue.choice.append(choice)
		
		dialogue_data.append(dialogue)