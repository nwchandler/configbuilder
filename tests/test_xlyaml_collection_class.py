import yaml
import xlyaml

def test_collection_class_buildOutput_builds_valid_yaml():
	sheet_object = [ \
		['key1', 'value1', None, None], \
		['key2', 'value2', None, None], \
		['list1', None,  None, None], \
		[None, 'l1_item1', None, None], \
		[None, 'l1_item2', None, None], \
		[' dict1', 'd1_key1', 'd1_key2', 'd1_key3'], \
		[None, 'd1_k1_v1', 'd1_k2_v1', 'd1_k3_v1'], \
		[None, 'd1_k1_v2', 'd1_k2_v2', 'd1_k3_v2'] \
		]
	collection_object = xlyaml.Collection(sheet_object)

	assert collection_object.collection == \
		yaml.load(collection_object.buildOutput())