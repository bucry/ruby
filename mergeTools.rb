#encoding:UTF-8

=begin
	@author : bfc
	@description : analysis office excel and create xml file
	@type 1
	@date : 2014-12-1
	@version : 1.0
=end

require'win32ole'
require "rexml/document" 
require 'pathname'

class EnginesManager
	$index = 0
	$SCRIPT_FILE_PATH = Pathname.new(__FILE__).realpath
	$SCRIPT_DIR_PATH = Dir.pwd;
	@@worksheetIndex = 3
	@@startRowIndex = 3
	@@startTitleIndex = 2
	
	def EnginesManager.start
		puts "toole initialize success"
	end
	
	def EnginesManager.end
		puts "toole stop success"
	end
end

class Engines < EnginesManager

	def createXml(fullData, nameDate, file)
		hashTable = Hash.new #create hash table
		#puts hashTable.include?("his")
			
		for findIsExitIndex in 1..fullData.length do
			if fullData[findIsExitIndex] != nil
				key = nil
				dataValue = [[],[]]
				for subFindIsExitIndex in 1..fullData[findIsExitIndex].length do
					if subFindIsExitIndex == 2
						if fullData[findIsExitIndex][subFindIsExitIndex] != nil
							key = fullData[findIsExitIndex][subFindIsExitIndex].to_s
						end
					end
				end
				
				if fullData[findIsExitIndex] != nil
					if hashTable.include?(key) == false
						dataValue[findIsExitIndex] = fullData[findIsExitIndex]
						hashTable[key] = dataValue
					else
						dataValue = hashTable[key]
						dataValue[findIsExitIndex] = fullData[findIsExitIndex]
						hashTable[key] = dataValue
					end
				end
			end
		end
		
		hashTable.each do |key, value|
			if key != nil
				if value != nil && value.length != 0 && key != nil
					puts key.encoding
					#puts "key:"+ key.encode("UTF-8") +",value："+value.encode("UTF-8")
				end
			end
		end
		#puts hashTable
		
		doc = REXML::Document.new 
		element = doc.add_element('root') #create root attribute
		
		hashTable.each do |key,value|
			if key != nil
				if value != nil && value.length != 0
					subElement = element.add_element("subRoot") #create root attribute
					titValue = nameDate[2].encode("UTF-8")
					titKey = key.encode("UTF-8")
					
					begin
						titKey = titKey.to_i
					rescue Exception
						titKey = titKey
					else
					ensure
					end
					
					if (titKey.is_a?(Float) == true)
						titKey = titKey.to_i
					else
						titKey = titKey
					end
					
					chapterRoot = subElement.add_element("item", {titValue=>titKey})
					for rootIndex in 1..value.length do
						if value[rootIndex] != nil && value[rootIndex].length != 0
							if value[rootIndex] != nil && value[rootIndex].length != 0
								chapter = chapterRoot.add_element("subItem") #create 
								begin
									for subFindIndex in 3..value[rootIndex].length do
										if nameDate[subFindIndex] != nil
											begin
												if value[rootIndex][subFindIndex] != nil
													if (value[rootIndex][subFindIndex].is_a?(Float) == true)
														va = value[rootIndex][subFindIndex].to_i
														chapter.add_attribute nameDate[subFindIndex], va
													else
														va = value[rootIndex][subFindIndex].to_s.encode("UTF-8")
														chapter.add_attribute nameDate[subFindIndex], va
													end
												elsif
													chapter.add_attribute nameDate[subFindIndex], value[rootIndex][subFindIndex]
												end
												
											rescue Exception
												puts "find error"
											else
												
											ensure
												
											end
										end
									end
								rescue Exception
									puts "add_attribute found Woring"
								else
									
								ensure
									puts "create Success"
								end
							end
						end
					end
				end
			end
		end
		file.puts doc.write #save file
	end
	
	#create ruby excel instance
	def analtsisExcel(excel)
		workbook = nil
		if (ARGV[0] == nil)
			begin
				puts "Please input excel file path:"
				filePath = gets
				workbook =excel.Workbooks.Open(filePath)
			rescue Exception
				puts "something is error, please try agin"
				exit()
			else
				puts "found file succee, start..."
			ensure
				if workbook == nil
					workbook.close
					excel.Quit #close excel stream
					exit()
				end
			end
		else
			workbook = excel.Workbooks.Open(ARGV[0])
		end
		return workbook
	end

	def analtsisExcelBath(excel, excelFilePath)
		workbook = nil
		begin
			workbook = excel.Workbooks.Open(excelFilePath)
			puts "find excel file localte : " + excelFilePath
		rescue => err
			puts err
			puts "find " + excelFilePath + " someThing is worong"
			exit
		else
			
		ensure
			if (workbook == nil)
				workbook.close
				excel.Quit #close excel stream
				exit()
			end
		end
		return workbook
	end
	
	
	#create xml file save path
	def createXmlFilePath
		file = nil
		if (ARGV[1] == nil)
			begin
				puts "Please input save xml file path:"
				xmlFilePath = gets
				file = File.new(xmlFilePath, "w+") 
			rescue Exception
				puts "something is error, please try agin"
				exit()
			else
				puts "found file succee, start..."
			ensure
				if file == nil
					exit();
				end
			end
		else
			file = File.new(ARGV[1], "w+")
		end
		return file
	end
	
	def createXmlFilePathBath(saveXmlFilePath)
		file = nil
		realSaveXmlFilePath = saveXmlFilePath.split("/")[saveXmlFilePath.split("/").length - 1]
		saveXmlFilePath = $SCRIPT_DIR_PATH + "/" + (realSaveXmlFilePath.split("-")[realSaveXmlFilePath.split("-").length - 1])
		begin
			file = File.new(saveXmlFilePath, "w+")
			#Iconv.iconv("GBK", "UTF-8", file)  
			puts "create xml file in " + saveXmlFilePath
		rescue Exception
			puts "find " + saveXmlFilePath + " someThing is worong"
			exit
		else
			
		ensure
			if (file == nil)
				exit
			end
		end
		return file
	end
	
	
	#init name arrays
	def initNameArray(column, worksheet)
		nameDate = []
		for nameIndex in 1..column do
			nameValue = worksheet.usedrange.cells(1, nameIndex).value
			puts nameValue
			if nameValue != nil
				nameDate[nameIndex] = nameValue.split("-")[1]
			end
		end 
		return nameDate
	end
	
	#init xml tile
	def initTitleArray(column, worksheet)
		titleDate = []
		for index in 1..column do
			titleDate[index] = worksheet.usedrange.cells(@@startTitleIndex, index).value
		end
		return titleDate
	end
	
	# init xml value
	def initFullData(row, column, worksheet)
		fullData = [[],[]]
		for i in @@startRowIndex..row do
			doc = REXML::Document.new 
			data = []
				for j in 1..column do
					begin
						data[j] = worksheet.usedrange.cells(i,j).value
					rescue Exception
						puts "found error"
					else
						
					ensure
						
					end
				end 
			fullData[i] = data;
		end
		return fullData
	end
	
	def closeExcelStream(workbook, excel)
		begin
			workbook.close
			excel.Quit #close excel stream
		rescue Exception
		else
			
		ensure
			
		end
	end
	
	def queryExcelColAndRow(worksheet)
		value = []
		worksheet.Select
		row = worksheet.usedrange.rows.count
		value[0] = row
		column = worksheet.usedrange.columns.count
		value[1] = column
		worksheet.usedrange.each{|cell|
			#puts cell.value
		}
		return value
	end
	
	def createWorksheet(workbook, index)
		return workbook.Worksheets(index) 
	end
	
	#start
	def start
		###read excel
		excel =WIN32OLE::new('excel.Application')
		workbook = analtsisExcel(excel)
		file = createXmlFilePath
		worksheet = createWorksheet(workbook, @@worksheetIndex)
		vo = queryExcelColAndRow(worksheet)
		#save attribute name to array
		nameDate = initNameArray(vo[1], worksheet)
		#save title name to array
		titleDate = initTitleArray(vo[1], worksheet)
		#save attribute value to array
		fullData = initFullData(vo[0], vo[1], worksheet)
		closeExcelStream(workbook, excel)
		createXml(fullData, nameDate, file)
	end
	
	
	def bathHandleExce(path)
		folderArray = iterationFolder(path)
		puts folderArray
		folderArray.each do |item|
			begin
				if ((item != nil) && (item.split(".")[1] == "xlsx" || item.split(".")[1] == "xls"))
					excel =WIN32OLE::new('excel.Application')
					workbook = analtsisExcelBath(excel, item)
					file = createXmlFilePathBath(item.split(".")[0].to_s + ".xml")
					worksheet = createWorksheet(workbook, @@worksheetIndex)
					vo = queryExcelColAndRow(worksheet)
					#save attribute name to array
					nameDate = initNameArray(vo[1], worksheet)
					#save title name to array
					titleDate = initTitleArray(vo[1], worksheet)
					#save attribute value to array
					fullData = initFullData(vo[0], vo[1], worksheet)
					closeExcelStream(workbook, excel)
					createXml(fullData, nameDate, file)
				else
					puts item.to_s + " not excel file"
				end
			rescue Exception
				closeExcelStream(workbook, excel)
				puts "find someThing is worong　" +item.force_encoding("UTF-8")
			else
				
			ensure
				
			end
		end
	end
	
	#iteration folder and find excel
	def iterationFolder(path) 
		folderArray = []
		Dir.entries(path).each do |sub|         
			if sub != '.' && sub != '..'  
			  if File.directory?("#{path}/#{sub}")  
				#puts "[#{sub}]"
				iterationFolder("#{path}/#{sub}")  
			  else  
				#puts "|--#{sub}"
				folderArray << $SCRIPT_DIR_PATH + "/" + "#{sub}"
			  end  
			end  
		end 
		return folderArray
	end 
	
end
en = Engines.new 
en.bathHandleExce($SCRIPT_DIR_PATH)