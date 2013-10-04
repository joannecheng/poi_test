class Cell
  def self.cell_by_type(poi_cell)
    case poi_cell.getCellType
    when cell.CELL_TYPE_NUMERIC
      NumericCell.new(poi_cell)
    when cell.CELL_TYPE_FORMULA
      FormulaCell.new(poi_cell)
    when cell.CELL_TYPE_BLANK
      BlankCell.new(poi_cell)
    end
  end

  def self.cell
    Java::OrgApachePoiSsUsermodel::Cell
  end

end

class NumericCell
  def initialize(cell)
    @cell = cell
  end
end

class FormulaCell
  def initialize(cell)
    @cell = cell
  end
end

class BlankCell
  def initialize(cell)
    @cell = cell
  end

end
