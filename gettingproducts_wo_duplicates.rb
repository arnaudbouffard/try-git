
require 'rubygems'
require 'nokogiri'
require 'open-uri'
require 'spreadsheet'

########################################################
# this script generates raw parsed results which need to
# be analysed for duplicates

allproducts_page_url = 'http://www.ikea.com/fr/fr/catalog/allproducts/department'
allproducts_nokonode = Nokogiri::HTML(open(allproducts_page_url))

# opening excel spreadsheet, naming sheets

Spreadsheet.client_encoding = 'UTF-8'
book = Spreadsheet::Workbook.new

sheet1 = book.create_worksheet
sheet1.name = 'categories'

sheet2 = book.create_worksheet
sheet2.name = 'subcategories'

sheet3 = book.create_worksheet
sheet3.name = 'products'
sht3_rows =  ['subcategories',
  'names',
  'refs',
  'descriptions',
  'prices',
  'currency and quantity',
  'url',
  'sold online']
  sht3_rows.each_with_index do |title, index|
    sheet3[0, index] = title
  end

  sheet4 = book.create_worksheet
  sheet4.name = 'analysis'

# writing product categories in first sheet

universes = allproducts_nokonode.css("span[class ='header']")
puts "universes length is #{universes.length}"
universes.each_with_index do |universe, index|
  sleep(rand(1))
  sheet1[index, 0] = "#{universe.text.strip!}"
end

# writing product subcategories in second sheet, with corresponding urls

subcategories = allproducts_nokonode.css('div.productCategoryContainerWrapper a')
i = 1

subcategories.each_with_index do |subcategory,ind|
  sleep(rand(1))
  next if ind < 81
  puts "now doing #{subcategory.text.strip!}, #{ind} subcategory out of
  #{subcategories.length}"
  sheet2[ind, 0] = "#{subcategory.text.strip!}"
  subcat_page_url = "http://www.ikea.com#{subcategory['href']}"
  sheet2[ind, 1] = subcat_page_url

# for each subcategory, parsing products and writing
# product details in 3rd sheet ...

subcat_nokonode = Nokogiri::HTML(open(subcat_page_url))

# subcat = subcat_nokonode.css('title')
# psubcat = subcat_nokonode.css("meta[name ='IRWStats.subCategory']")
# puts "page title as read in title tag is #{subcat.text}"
# puts "page subcat as read in meta tag is #{psubcat}"
# puts "page subcat as source link is #{subcategory.text.strip!}"

# product subcategory and name

names = subcat_nokonode.css("div[class ='productTitle floatLeft']")
names.each_with_index do |name, index|
  # puts "currently parsing #{name} product"

  # puts "putting product subcat in cell(#{i + index}, 0)"
  sheet3[i + index, 0] = subcategory.text.strip!

  # puts "putting #{i}nth product name in cell(#{i + index}, 1)"
  sheet3[i + index, 1] = "#{name.text}"
end

  # product description

  descriptions = subcat_nokonode.css("div[class ='productDesp']")
  descriptions.each_with_index do |description, index|
    # puts "putting #{subcat.text[0..-8]} #{index}nth product description
    # in cell(#{i + index}, 3)"
    sheet3[i + index, 3] = "#{description.text}"
  end

  # product price, currency and quantity

  prices = subcat_nokonode.css("div[class ='price regularPrice']")
  prices.each_with_index do |price, index|
    # puts "putting #{index}nth product price in cell(#{i + index}, 4)"
    fullprice = price.text[0...35].strip!
    sheet3[i + index, 4] = fullprice.slice(0..(fullprice.index('€') - 2))
    sheet3[i + index, 5] = fullprice.slice((fullprice.index('€'))..-1)
  end

  # products urls, refs, and availability online

  products_pages_end_of_urls = subcat_nokonode.css("a[class ='productLink']")
  products_pages_end_of_urls.each_with_index do |products_pages_end_of_url, index|
    product_page_url = "http://www.ikea.com#{products_pages_end_of_url['href']}"
    # puts "putting #{index}nth product url in cell(#{i + index},<6></6>)"
    sheet3[i + index, 6] = product_page_url

    begin
      file = open(product_page_url)
      product_nokonode = Nokogiri::HTML(file)
      ref = product_nokonode.css("div[id ='itemNumber']")
      # puts ref.length
      # puts "putting #{subcat.text[0..-8]} #{index}nth product ref
      # in cell(#{i + index}, 2)"
      # puts "ref is #{ref.text}"
      sheet3[i + index, 2] = "#{ref.text}"
      a = product_nokonode.css("div[class ='buttonContainer'] input[value = 'Acheter
        en ligne']")
      sheet3[i + index, 7] = !a.empty?
    rescue
      puts "invalid URL"
      sheet3[i + index, 2] = "invalid URL"
      sheet3[i + index, 7] = "invalid URL"
    end
  end

  i += names.length
  book.write '/users/arnaudbouffard/complete_search.xls'

end

sheet4[0, 0] = 'there are'
sheet4[0, 1] = i
sheet4[0, 2] = 'products in the online catalog'

sheet4[1, 0] = 'there are'
sheet4[1, 1] = subcategories.length
sheet4[1, 2] = 'subcategories'

book.write '/users/arnaudbouffard/complete_search.xls'

# check out  http://nokogiri.org/Nokogiri/XML/NodeSet.html

# the markup for a link is: <a id="txt31" href="/fr/fr/catalog/categories/
# departments/workspaces/16195/">Organisation des cables et accessoires</a>
# parse all the links (<a> tags) inside the CSS
# class '.productCategoryContainerWrapper'
# check Nokogiri documentation for parsing innerHtml and Href
# for each of them, create a new Category object: Category.new(name, url)

# ok so now you have objects like this
# #<Category @name="Bureaux et bureaux pour ordinateur" @url="/fr/fr/catalog/
# categories/departments/workspaces/16195/">

# Then, for each Category page, visit the url
# Parse the css classes ".product"
# it seems the parent tag of each product is <div id="item_410148622_1"
# class="threeColumn product " title="">
# for each product, create a new Product object

