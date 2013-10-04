require 'poi'

filename = 'Example Model.xls'
p = Poi.new(filename)

input = p.worksheet(0)
summary = p.worksheet(1)
cons = p.worksheet(5)

turbine_construction = input.row(23).cell(4)
puts turbine_construction.value
summary_total = summary.row(17).cell(5)
puts summary_total.value
p.evaluator.evaluate(summary_total.cell)
