module Pexcel
	class Espreadsheet
		def self.create
			excel_obj = Spreadsheet::Workbook.new # We have created a new object of the Spreadsheet book

			#sheet = book.create_worksheet(name: 'First sheet') # We are creating new sheet in the Spreadsheet(We can create multiple sheets in one Spreadsheet book)
      sheet = excel_obj.create_worksheet :name => "New Work Sheet"
      header_row_count = 0
      bold = Spreadsheet::Format.new :weight => :bold
      border = Spreadsheet::Format.new :top => :thin,
                                       :right => :thin,
                                       :left => :thin,
                                       :bottom => :thin
      bold_border = Spreadsheet::Format.new :weight => :bold,
                                            :top => :thin,
                                            :right => :thin,
                                            :left => :thin,
                                            :bottom => :thin
      4.times do |x| sheet.row(header_row_count).set_format(x , bold) end
      	4.times do |x| sheet.row(header_row_count).set_format(x , bold_border) end
			# syntax to create new row is as the following:
			# sheet.row(row_number).push(column first', 'column second', 'column third')

			# Let's create first row as the following.
			sheet.row(0).push('Test Name', 'Test country', 'Test city', 'Test profession') # Number of arguments will be number of columns

			# We can create many rows same as the mentioned above.
			sheet.row(1).push('Bobby', 'US', 'New York', 'Doctor')
			sheet.row(2).push('John', 'England', 'Manchester', 'Engineer')
			sheet.row(3).push('Rahul', 'India', 'Mumbai', 'Teacher')

			# Write this sheet's contain to the test.xls file.
			excel_obj.write Rails.root.join('app', 'assets', 'images', 'text.xls')
			
		end


      def self.generate_pdf_for_abc
      	dirname = Rails.root.join('app', 'assets', 'images')
        Prawn::Font::AFM.hide_m17n_warning = true
        pdf = Prawn::Document.new(page_size: 'A4', page_layout: :landscape)
        pdf.font_families.update("Arial" => {
            :normal => Rails.root.join("app/assets/fonts/OpenSans-Regular.ttf"),
            :italic => Rails.root.join("app/assets/fonts/OpenSans-Regular.ttf"),
            :bold => Rails.root.join("app/assets/fonts/OpenSans-Bold.ttf"),
            :bold_italic => Rails.root.join("app/assets/fonts/OpenSans-Bold.ttf")
        })
        pdf.font "Arial"
        payment_date = Date.today.to_date
  
        two_dimensional_array = pdf.make_table([['COMMEASURE SOLUTIONS PHILIPPINES INC', size: 15], ["LG1 Cityland III 105 VA Rufino St. Legaspi Village, San Lorenzo, Makati City, NCR 1223", align: :center], [""], ["VAT REG. TIN: 009-654-476-000", size: 15]], :cell_style => {:borders => [], :padding => [2, 2, 2, 2], :align => :center, :text_color => "0070c0", :font_style => :bold})
        ak_r = pdf.make_table([["OFFICIAL RECEIPT"], ["(Other Information Technology and Computer Service Activities)"]], :cell_style => {:borders => [], :padding => [0, 0, 0, 0], :align => :left, :font_style => :bold})
        booking_data = [["Booking ID", "Guest Name", "Booking Source"]]
        
        rev_row = [['TRDN xxx ', '  Rev. No. xxx ', ' Rev. Date xx/xx/xxx']]
        pdf.table(rev_row, :cell_style => {:borders => [], :padding => [0, 20, 10, 0]}) do
          column(0).style :width => 200
          column(1).style :width => 200
          column(3).style :width => 200
        end
        pdf.move_down(10)
          
        #booking_details = pdf.make_table(booking_data, :cell_style => {:borders => [], :padding => [5, 5, 5, 5]})

        pdf.font_size(16) do
          pdf.text "CLIENT’S COPY", align: :center, font_style: :bold
          pdf.fill_color "f70d1a"
          pdf.text "ACCN: AC_RDO_MMYYYY_XXXXXX", font_style: :bold
        end
        ptu_table = [
          ["PTU No. ", "AC-xxx-xxxxxx-xxxxxx", "Sup Name: Commeasure Solutions Philippines, Inc"],
          ["PTU No. Date Issued", "xx/xx/xxxx", "Sup Address: LG-1 Cityland III 105 V.A. Rufino Street, Legaspi Village, San Lorenzo City, Makati, NCR, Fourth District, Philippines 1223"],
          ["PTU No. Valid Until", "xx/xx/xxxx",  "Sup TIN: 009-654-476-00000" ],
          ["Document No. Range: ", "", "Acc. No.:"],
          ["From: ","", "Acc. No. Date Issued:"],
          ["To: ", "", "Acc. No. Valid Until:" ]
        ]
        pdf.move_down(10)
        pdf.table(ptu_table, :cell_style => {:borders => [], :padding => [0, 5, 5, 5], text_color: '0070c0'}) do
          
        end
        pdf.move_down(10)

        software_version = [["Software Version: MidOffice (Rails 5.2.6, Ruby 2.6.10)."]]
        pdf.table(software_version, :cell_style => {:borders => [], :padding => [0, 5, 5, 5], text_color: '0070c0' , :size => 8, :font_style => :bold}) do
          
        end

        pdf.move_down(20)

        doc_five_year = [["This receipt file is generated & sent to the client on #{Date.today.strftime("%B")} #{Date.today.day} #{Date.today.strftime("%Y")}."]]
        pdf.table(doc_five_year, :cell_style => {:borders => [], :padding => [0, 5, 5, 5], text_color: '0070c0' , :size => 8}) do
          
        end   
        pdf.render_file "#{dirname}/RDMONTH-#{Date.today.strftime("%m%y")}.pdf"
      end

		# def read_spreadsheet(file_path)
		#   workbook = Spreadsheet.open(file_path)
		#   worksheet = workbook.worksheet(0) # assuming there's only one worksheet

		#   rows = []
		#   worksheet.each do |row|
		#     rows << row.map(&:to_s)
		#   end

		#   return rows
		# end

		# def self.write_to_spreadsheet(data, file_path)
		#   workbook = Spreadsheet::Workbook.new
		#   worksheet = workbook.create_worksheet

		#   data.each_with_index do |row_data, row_index|
		#     row_data.each_with_index do |cell_data, col_index|
		#       worksheet[row_index, col_index] = cell_data
		#     end
		#   end

		#   workbook.write(file_path)
		# end

	 #  def self.write
	 #    data = [
	 #      ["Name", "Age", "Email"],
	 #      ["John Doe", 30, "john@example.com"],
	 #      ["Jane Smith", 25, "jane@example.com"]
	 #    ]
	 #    file_path = Rails.root.join('app', 'assets', 'images', 'file.xls')
	 #    write_to_spreadsheet(data, file_path)
	 #  end
	end
end