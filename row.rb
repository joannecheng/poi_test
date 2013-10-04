class Row
  attr_reader :row
  def initialize(row)
    @row = row
  end

  def cells
    cell_count.times.map do |i|
      cell(i)
    end
  end

  def cell(index)
    cell = @row.getCell(index)
    Cell.cell_by_type(cell)
  end

  private

  def cell_count
    @row.getLastCellNum
  end

end

