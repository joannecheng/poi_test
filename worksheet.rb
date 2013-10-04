class Worksheet
  attr_reader :worksheet
  def initialize(worksheet)
    @worksheet = worksheet
  end

  def rows
    row_count.times.map do |i|
      row(i)
    end
  end

  def row(index)
    Row.new @worksheet.getRow(index)
  end

  def name
    @worksheet.getSheetName()
  end

  private

  def row_count
    @worksheet.getLastRowNum
  end

end

