
class Sku:
    def __init__(self, color, size, productNo, quantity=0, image=None, imageFile=None):
        self.color = color
        self.size = size.upper()
        self.productNo = productNo
        self.quantity = 0
        self.image = None
        self.imageFile = None
        if self.size != 'S' and self.size != 'M' and self.size != 'L' \
                and self.size != 'XL' and self.size != '2XL' \
                and self.size != '3XL' and self.size != '4XL' \
                and self.size != '5XL':
            raise ValueError("尺码不规范")

    def create_sku(skuStr):
        trimmed_skuStr = skuStr.strip()  # 对skuStr进行trim处理
        if '-' in trimmed_skuStr:
            parts = trimmed_skuStr.split('-')
            num_parts = len(parts)

            if num_parts == 2:
                productNo = parts[0]
                size = parts[1]
                return Sku("---", size, productNo)
            elif num_parts == 3:
                productNo = parts[0]
                color = parts[1]
                size = parts[2]
                return Sku(color, size, productNo)
            else:
                raise ValueError("sku格式不规范")


# 示例用法
# sku1 = Sku.create_sku("ABC123-XL")
# print(sku1.productNo)  # 输出: ABC123
# print(sku1.size)       # 输出: XL
# print(sku1.color)      # 输出: 空字符串
#
# sku2 = Sku.create_sku("ABC123-Red-XL")
# print(sku2.productNo)  # 输出: ABC123
# print(sku2.size)       # 输出: XL
# print(sku2.color)      # 输出: Red
# print(sku2.quantity)      # 输出: Red

# sku3 = Sku.create_sku("InvalidSKUFormat")
# 输出: ValueError: sku格式不规范
