groups = Array.new(4,Array.new)
p groups
4.times do |i|
  arr = Array.new
  arr[0] = i
  arr[1] = "er"
  groups[i].push arr
end
p groups
