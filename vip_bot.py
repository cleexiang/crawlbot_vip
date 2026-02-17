import asyncio
from playwright.async_api import async_playwright
import openpyxl
from openpyxl.styles import Font
from datetime import datetime
import os
import random

get_page_number_js = r"""
() => {
  // 查找包含总页数的元素
  const span = document.querySelector('span.total-item-nums');
  if (!span) return null;
  // 提取出“共27页”这种文本里的数字部分
  const match = span.textContent.match(/共(\d+)页/);
  return match ? parseInt(match[1], 10) : null;
}
"""


async def get_page_number(keyword):
    """
    获取某个商品的页数
    """
    async with async_playwright() as p:
        browser = await p.chromium.connect_over_cdp("http://localhost:9222")
        contexts = browser.contexts
        if contexts:
            page = contexts[0].pages[0]
            await page.goto(f"https://category.vip.com/suggest.php?keyword={keyword}")
            print(f"Current URL: {page.url}")
            page_number = await page.evaluate(get_page_number_js)
            print(f"共：{page_number} 页")
            return page_number


get_items_of_page_js = r"""
() => {
  // 获取所有商品元素
  const items = Array.from(document.querySelectorAll('div.c-goods-item.J-goods-item.c-goods-item--auto-width[data-product-id]'));
  return items.map(item => {
    // 商品ID
    const productId = item.getAttribute('data-product-id');

    // 商品链接
    const aTag = item.querySelector('a[href*="detail"]');
    // 有部分链接是 // 开头，加上协议
    let href = aTag ? aTag.getAttribute('href') : null;
    if (href && href.startsWith("//")) {
      href = "https:" + href;
    }

    // 商品底部信息块
    const bottom = item.querySelector('div.c-goods-item-bottom');
    if (!bottom) return null;

    // 主要价格（特卖价）
    const salePriceDiv = bottom.querySelector('div.c-goods-item__sale-price');
    let salePrice = null;
    if (salePriceDiv) {
      const priceText = salePriceDiv.textContent.replace(/[^\d.]/g, '');
      salePrice = priceText ? parseFloat(priceText) : null;
    }

    // 原价（市场价）
    const marketPriceDiv = bottom.querySelector('div.c-goods-item__market-price');
    let marketPrice = null;
    if (marketPriceDiv) {
      const priceText = marketPriceDiv.textContent.replace(/[^\d.]/g, '');
      marketPrice = priceText ? parseFloat(priceText) : null;
    }

    // 折扣
    const discountDiv = bottom.querySelector('div.c-goods-item__discount');
    let discount = null;
    if (discountDiv) {
      discount = discountDiv.textContent.trim();
    }

    // 商品名称
    const nameDiv = bottom.querySelector('div.c-goods-item__name');
    let name = null;
    if (nameDiv) {
      name = nameDiv.textContent.trim();
    }

    return {
      productId,
      href,
      salePrice,
      marketPrice,
      discount,
      name
    };
  }).filter(x => x !== null);
}
"""

get_detail_info_js = r"""
() => {
  const result = {
    sizes: [],
    productCode: ''
  };

  // 抓取所有尺码
  const sizeElements = document.querySelectorAll('span.size-list-item-name');
  if (sizeElements.length > 0) {
    result.sizes = Array.from(sizeElements).map(el => el.textContent.trim());
  }

  // 抓取商品编码 - 方法1：在tbody中查找
  const tbody = document.querySelector('tbody.J_dc_table');
  if (tbody) {
    const thElements = tbody.querySelectorAll('th.dc-table-tit');
    for (let th of thElements) {
      if (th.textContent.includes('商品编码')) {
        // 找到最近的td兄弟元素
        let nextEl = th.nextElementSibling;
        while (nextEl && nextEl.tagName !== 'TD') {
          nextEl = nextEl.nextElementSibling;
        }
        if (nextEl) {
          result.productCode = nextEl.textContent.trim();
          break;
        }
      }
    }
  }

  // 如果还没有找到，尝试方法2：查找整个页面中包含"商品编码"的元素
  if (!result.productCode) {
    const allText = document.body.innerText;
    const match = allText.match(/商品编码[：:]\s*([A-Z0-9]+)/);
    if (match) {
      result.productCode = match[1];
    }
  }

  return result;
}
"""

def write_items_to_excel(items, keyword, output_file='products.xlsx'):
    """
    将items数据写入Excel文件，每个尺码对应一行数据
    """
    # 确保输出目录存在
    output_dir = os.path.dirname(output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 检查文件是否存在，如果存在则加载，否则创建新文件
    if os.path.exists(output_file):
        wb = openpyxl.load_workbook(output_file)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "商品数据"

        # 设置表头
        headers = ["标题", "市场价", "尺码", "商品编码", "折扣", "品牌", "产品id", "详情页"]
        ws.append(headers)

        # 设置表头样式
        header_font = Font(bold=True)
        for cell in ws[1]:
            cell.font = header_font

    # 写入数据 - 每个尺码对应一行
    total_rows = 0
    for item in items:
        # 获取尺码列表和商品编码
        sizes = item.get('sizes', [])
        product_code = item.get('productCode', '')

        # 如果没有尺码，则创建一条记录（尺码为空）
        if not sizes:
            sizes = ['']

        # 为每个尺码创建一行
        for size in sizes:
            row = [
                item.get('name', ''),           # 标题
                item.get('salePrice', ''),    # 市场价
                size,                           # 尺码
                product_code,                   # 商品编码
                item.get('discount', ''),       # 折扣
                keyword,                        # 品牌 (keyword)
                item.get('productId', ''),      # 产品id
                item.get('href', '')            # 详情页链接
            ]
            ws.append(row)
            total_rows += 1

    # 自动调整列宽
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width

    # 保存文件
    wb.save(output_file)
    print(f"✓ Excel文件已更新: {output_file}")
    print(f"✓ 共导入 {total_rows} 条商品数据")


async def human_scroll(page):
    """
    模拟真实用户的滑动，每次滑动到页面最底部，然后等待3秒加载内容
    连续2次页面高度不变则停止滚动
    """
    last_height = 0
    no_change_count = 0

    while no_change_count < 2:  # 连续2次高度不变则停止滚动
        # 滑动到页面最底部
        await page.evaluate('window.scrollTo(0, document.documentElement.scrollHeight)')
        print(f"  已滑动到页面底部")

        # 等待3秒让内容加载
        await asyncio.sleep(3)

        # 获取当前页面高度
        current_height = await page.evaluate('document.documentElement.scrollHeight')

        if current_height == last_height:
            no_change_count += 1
            print(f"  页面高度未变化 ({no_change_count}/2) - 高度: {current_height}px")
        else:
            no_change_count = 0
            print(f"  页面高度已变化 - 上次: {last_height}px, 当前: {current_height}px")

        last_height = current_height

    print("✓ 页面滚动完成，无更多内容加载")


async def get_detail_info(page_obj, item):
    """
    打开详情页并获取尺码和商品编码信息
    """
    try:
        if item.get('href'):
            await page_obj.goto(item['href'], wait_until='networkidle')
            detail_info = await page_obj.evaluate(get_detail_info_js)
            item['sizes'] = detail_info.get('sizes', [])
            item['productCode'] = detail_info.get('productCode', '')
            print(f"✓ 获取详情: {item['name'][:20]} - 尺码数: {len(item['sizes'])} - 编码: {item['productCode']}")
        return item
    except Exception as e:
        print(f"✗ 获取详情失败: {item['name'][:20]} - {str(e)}")
        return item


async def get_items_of_page(keyword, page):
    """
    获取某某个商品某页的数据，并获取每个商品的详情信息
    """
    async with async_playwright() as p:
        browser = await p.chromium.connect_over_cdp("http://localhost:9222")
        contexts = browser.contexts
        if contexts:
            page_obj = contexts[0].pages[0]
            await page_obj.goto(f"https://category.vip.com/suggest.php?keyword={keyword}&page={page}")
            print(f"Current URL: {page_obj.url}")

            # 模拟用户滑动，加载所有商品数据
            print("正在加载页面数据，请稍候...")
            await human_scroll(page_obj)

            items = await page_obj.evaluate(get_items_of_page_js)
            print(f"✓ 获取到 {len(items)} 个商品")

            # 逐个打开详情页获取尺码和商品编码
            print("开始获取商品详情...")
            for item in items:
                await get_detail_info(page_obj, item)

            # 获取数据后写入Excel，文件名为: keyword_YYYY-MM-DD.xlsx
            if items:
                today = datetime.today().strftime('%Y-%m-%d')
                output_file = f"data/{keyword}_{today}.xlsx"
                write_items_to_excel(items, keyword, output_file)
            
async def main():
    keyword = "阿迪达斯"

    # 获取总页数
    total_page = await get_page_number(keyword=keyword)

    if total_page:
        print(f"\n开始抓取 {keyword} 的商品数据，共 {total_page} 页")
        print("=" * 50)

        # 遍历每一页
        for page in range(1, total_page + 1):
            print(f"\n【第 {page}/{total_page} 页】")
            await get_items_of_page(keyword=keyword, page=page)

        print("\n" + "=" * 50)
        print(f"✓ {keyword} 的所有数据抓取完成！")


if __name__ == "__main__":
    # asyncio.run()
    asyncio.run(main())