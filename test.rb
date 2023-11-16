require "google_drive"

session = GoogleDrive::Session.from_config("config.json")

ws1 = session.spreadsheet_by_key("1Y68QzhD92oWOp4Sxg4bZqMQDlIETz6a-GENTj5ex_PU").worksheets[0]
ws2 = session.spreadsheet_by_key("1Y68QzhD92oWOp4Sxg4bZqMQDlIETz6a-GENTj5ex_PU").worksheets[1]

class Tabela
	include Enumerable

	def initialize(worksheet)
		@worksheet = worksheet
		@tabela = []
		@imenaKolona = []
		@indeksi = {}
		@pocetni_red = nil
		@pocetna_kolona = nil
		@prazni_redovi = []
		@duzina_u_redovia
		@sirina_u_redovim
	end

	attr_accessor :worksheet, :tabela, :imenaKolona, :indeksi, :pocetni_red,
	:pocetna_kolona, :prazni_redovi, :duzina_u_redovima, :sirina_u_redovima


	def createWorksheet(worksheet)
		@worksheet = worksheet

		(1..worksheet.num_rows).each do |row|
			break unless @pocetni_red.nil? && @pocetna_kolona.nil?

			(1..worksheet.num_cols).each do |col|
				cell_value = worksheet[row, col]

				if cell_value != '' && @pocetni_red.nil? && @pocetna_kolona.nil?
					@pocetni_red = row
					@pocetna_kolona = col
					@duzina_u_redovima = worksheet.num_rows - @pocetni_red + 1
					@sirina_u_kolonama = worksheet.num_cols - @pocetna_kolona + 1
					break
				end
			end
		end
	end

	def getHeader(worksheet)
		@worksheet = worksheet

		(@pocetna_kolona..worksheet.num_cols).each do |col|
			@imenaKolona.push(worksheet[@pocetni_red, col])
		end
	end

	def popuniTabelu(worksheet)
		pomocni_niz = []

		(2..@duzina_u_redovima).each do |row|
			(1..@sirina_u_kolonama).each do |col|
				cell_value = worksheet[@pocetni_red + row - 1, @pocetna_kolona + col - 1]

				if cell_value == '' || cell_value.downcase == 'total' || cell_value.downcase == 'subtotal'
					@prazni_redovi << row
					break
				else
					pomocni_niz << cell_value
				end
			end
			@tabela << pomocni_niz unless pomocni_niz.empty?
			pomocni_niz = []
		end
	end
	
	def uzmiIndekse
		indeks = {}

		@imenaKolona.each_with_index do |ime_kolone, index|
			indeks[ime_kolone] = OurColumn.new(self, index)
		end

		indeks
	end

	def row(broj)
		@tabela[broj] if broj >= 0 && broj < @tabela.length
	end

	def each
		@tabela.each do |row|
			row.each do |column|
				yield column if block_given?
			end
		end
	end

	def [](kolona)
		indeksi[kolona]
	end

	def +(drugaTabela)
		if @imenaKolona == drugaTabela.imenaKolona
			novi_red = @worksheet.num_rows + 1

			drugaTabela.tabela.each do |red|
				@imenaKolona.each_with_index do |ime_kolone, indeks_kolone|
					@worksheet[novi_red, indeks_kolone + 2] = red[indeks_kolone]
				end
				novi_red += 1
			end

			@worksheet.save
			@worksheet.reload
		end
	end
end

class Kolona 

	def initialize(tabela,kolona)
		@tabela = tabela                
		@kolona = kolona               
	end

	attr_accessor :tabela, :kolona


	def nadjiIndeks
		@tabela.tabela.each_index do |index|
			return index if @tabela.tabela[index][@kolona_index] == self
		end
		nil
	end

	def [](vrednost)
		@tabela.tabela[vrednost][@kolona_index]
	end

	def[]=(indeks,vrednost)
		@tabela.tabela[indeks][@kolona_index] = vrednost
	end
end

tabela_instance = Tabela.new(ws1)
tabela_instance.createWorksheet(ws1)

tabela_instance.getHeader(ws1)
p tabela_instance.imenaKolona

tabela_instance.popuniTabelu(ws1)

p "vraca dvodimenzioni"
p tabela_instance.tabela

p ".row(0)"
p tabela_instance.row(0)

p ".each"
tabela_instance.each do |cell|
	puts cell
end

# p tabela_instance["Prva kolona"]
# p tabela_instance["Prva kolona"][7]
# tabela_instance["Prva kolona"][7] = "promena"
# p tabela_instance["Prva kolona"][7]
# tabela_instance.posaljiKolonuNaServer("Prva kolona", 7)
# p tabela_instance["Prva kolona"][7] 

tabela_instance2 = Tabela.new(ws2)
tabela_instance2.createWorksheet(ws2)


tabela_instance2.getHeader(ws2)
p tabela_instance2.imenaKolona

tabela_instance2.popuniTabelu(ws2)
p tabela_instance2.tabela


tabela_instance + tabela_instance2
