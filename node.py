import re


class TreeNode(object):
    def __init__(self, data):
        self.data = data
        self.children = []


class MultiTree(object):
    def __init__(self, data):
        self.tree = TreeNode(data)
        self.node_list = []

    def clear(self):
        self.node_list.clear()

    def get(self):
        return self.node_list

    def add(self, node, parent, tree):
        if parent == tree.data.get('id') and node not in tree.children:
            tree.children.append(node)
            return
        for child in tree.children:
            if child.data.get('id') == parent and node not in child.children:
                # children_list = child.children
                # children_list.append(node)
                # child.children = children_list
                child.children.extend([node])
                break
            else:
                self.add(parent, node, child)

    def traverse(self, tree):
        if isinstance(tree, MultiTree):
            # print(tree.tree.data.get('id'), tree.tree.data.get('pid'), tree.tree.data.get('code'),
            #       tree.tree.data.get('name'))
            if len(tree.tree.children) == 0:
                return
            self.traverse(tree.tree)
        else:
            children = []
            for child in tree.children:
                children.append(child)
                # print(child.data.get('id'), child.data.get('pid'), child.data.get('code'), child.data.get('name'))
            self.node_list.extend(children)
            while len(children) > 0:
                self.traverse(children.pop(0))

    def search_tree(self, name, count, depth, tree):
        if isinstance(tree, MultiTree):
            if count == depth and tree.tree.get('name') in name:
                print(tree.tree.get('code'))
                return tree.tree.get('code')
            if len(tree.tree.children) == 0:
                return
            self.search_tree(name, count + 1, depth, tree.tree)
        else:
            children = []
            for child in tree.children:
                children.append(child)
                if count == depth and child.data.get('name') in name:
                    print(child.data.get('code'))
                    return child.data.get('code')
            while len(children) > 0:
                self.search_tree(name, count + 1, depth, children.pop(0))

    # def search_name(self, name, tree):
    #     if isinstance(tree, MultiTree):
    #         if tree.tree.get('name') in name:
    #             return tree.tree.get('code')
    #         if len(tree.tree.children) == 0:
    #             return None
    #         self.search_name(name, tree.tree)
    #     else:
    #         if tree.data.get("name") in name:
    #             print(tree.data.get("code"), tree.data.get("name"))
    #             return tree.data.get("code")
    #
    #         children = []
    #         for child in tree.children:
    #             children.append(child)
    #         while len(children) > 0:
    #             self.search_name(name, children.pop())

    def prepare_search_name(self, tree):
        self.node_list.clear()
        self.traverse(tree)

    def search_name(self, name):
        res = []
        for node in self.node_list:
            if self.remove_char(name) == self.remove_char(node.data.get("name")):
                res.append(node.data.get("code"))
        return res

    @staticmethod
    def remove_char(string):
        return "".join(re.findall(r'[\u4e00-\u9fff]+', string))


if __name__ == '__main__':
    print()
    # sample = '国内增值税增值税'
    # for n in re.findall(r'[\u4e00-\u9fff]+', sample):
    #     print(n)
    # print('_' * 60)
    # print("".join(re.findall(r'[\u4e00-\u9fff]+', sample)))
    print("".join(re.findall(r'[\u4e00-\u9fff]+', '国有土地收益基金收入▲')))
    print(len("".join(re.findall(r'[\u4e00-\u9fff]+', '国有土地收益基金收入▲'))))