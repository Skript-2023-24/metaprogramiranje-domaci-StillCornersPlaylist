require 'google_drive'

def from_camel_case(str)
  str.to_s.gsub(/([A-Z])/, ' \1').strip.downcase
end

class Table
  attr_reader :start_col, :start_row

  def initialize
    session = GoogleDrive::Session.from_config("config.json")
    @ws = session.spreadsheet_by_key('1752Q5uKm5vbCU49faGWieQX1zIkrJ9bJZkoMp9yCIJ0').worksheets[0]
    @start_col = 0
    @start_row = 0
    find_header
  end

  def find_header
    t = table
    t.each_with_index do |row, row_index|
      row.each_with_index do |cell, index|
        if t[row_index][index] != ''
          @start_col = index
          @start_row = row_index
          break
        end
      end
    end
  end

  def table
    @ws.rows
  end

  def update_cell(row, col, new_value)
    @ws[row, col] = new_value
    @ws.save
  end

  def row(index)
    table[index]
  end

  def column(index)
    table.map { |row| row[index] }
  end

  def cells
    table.flatten
  end

  def [](key)
    list = []
    r = row(start_row)
    r.each_with_index do |cell, index|
      if cell == key
        list = column(index)
        break
      end
    end
    list
  end

  def []=(key, value)
    column_name = key.to_s.capitalize
    index = row(start_row).index(column_name)

    if index
      column_values = table.transpose[index]
      column_values[start_row] = value
      column_values.each_with_index do |val, row_index|
        @ws[start_row + row_index, index + @start_col] = val
      end
      @ws.save
    end
  end

  def method_missing(method_name, *args, &block)
    if method_name.to_s.end_with?('=')
      set_column_value(method_name.to_s.chomp('=').to_sym, args.first)
    else
      column_name = from_camel_case(method_name)
      row(start_row).each_with_index do |cell, index|
        return NewArray.new(column(index)) if cell.downcase == column_name
      end
      super
    end
  end

  def each(&block)
    table.each(&block)
  end

  def remove_total_subtotal_rows
    @ws.rows.reject! { |row| total_or_subtotal_row?(row) }
    @ws.save
  end

  def +(other_table)
    raise 'Tables have different headers' unless table_headers == other_table.table_headers

    @ws.rows += other_table.table
    @ws.save

    self
  end

  def -(other_table)
    raise 'Tables have different headers' unless table_headers == other_table.table_headers

    @ws.rows.reject! { |row| other_table.table.include?(row) }
    @ws.save

    self
  end

  def table_headers
    table[start_row]
  end

  private

  def total_or_subtotal_row?(row)
    row.any? { |cell| cell.downcase.include?('total') || cell.downcase.include?('subtotal') }
  end

  def set_column_value(column_name, value)
    index = row(start_row).index(column_name.to_s.capitalize)

    if index
      column_values = table.transpose[index]
      column_values[start_row] = value

      column_values.each_with_index do |val, row_index|
        @ws[start_row + row_index, index + @start_col] = val
      end

      @ws.save
    end
  end
end

class NewArray < Array
  def initialize(*args)
    super(args.flatten)
  end

  def sum
    inject(:+)
  end

  def avg
    sum.to_f / size
  end

  def method_missing(method_name, *args, &block)
    row_name = from_camel_case(method_name)
    each_with_index do |cell, index|
      return index if cell.downcase == row_name
    end
    super
  end
end

t = Table.new
puts t.table
puts "\n"
puts t.prva_kolona.test
