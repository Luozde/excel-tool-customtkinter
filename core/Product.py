class Product:
    def __init__(self, key2_list, key1):
        self.key2_list = key2_list
        self.key1 = key1


class Detail:
    def __init__(self, image, imageFile, color, sex, text, s, m, l, l1, l2, l3, l4, l5):
        self.image = image
        self.imageFile = imageFile
        self.color = color
        self.sex = sex
        self.text = text
        self.s = s
        self.m = m
        self.l = l
        self.l1 = l1
        self.l2 = l2
        self.l3 = l3
        self.l4 = l4
        self.l5 = l5


class Key2:
    def __init__(self, key2, details):
        self.key2 = key2
        self.details = details
