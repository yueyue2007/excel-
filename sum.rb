require 'roo'

class User
  def initialize
    @name  = ""
    @scores_dabiao = 0
    @scores_dati = 0
    @scores = 0
    @dabiao = []
    @dati = []
    @dui = ""
  end
  attr_accessor :scores, :dabiao, :dati ,:scores_dabiao, :scores_dati,:dui,:name
end
# 打开abc.xlsx，统计
xlsx = Roo::Spreadsheet.open('./abc2.xlsx')

people = Hash.new
xlsx.each(uid:'uid',name:'name',type:'type',description:'description',score:'score',time:'时间') do |item|
   if item[:uid] == 'uid' then
     next
   end
   if item[:time] < DateTime.new(2016,10,18,0,0) then
     next
   end
  #p item[:time]
  #p Time.now
  #p people[item[:name]] unless people[item[:name]]

  #people[item[:name]] = User.new unless people[item[:name]]
  unless people[item[:uid]] then
    people[item[:uid]] = User.new
    people[item[:uid]].name = item[:name]
  end
  if item[:type] == 2 then #每日达标
    people[item[:uid]].scores_dabiao += item[:score].to_i
    people[item[:uid]].dabiao.push item[:time].day
  elsif item[:type] == 7 then #每日答题
    people[item[:uid]].scores_dati += item[:score].to_i
    people[item[:uid]].dati.push item[:time].day
  end
  people[item[:uid]].scores =  people[item[:uid]].scores_dabiao +  people[item[:uid]].scores_dati
end



fenzu = Roo::Spreadsheet.open('./fenzu.xlsx')
#puts fenzu.info


fenzu.each(id:'ID',name:'姓名',dui:'部门') do |item|
  if item[:name] == '姓名' then
    next
  end
  #p item[:dui]
  people[item[:id]].dui  = item[:dui ]
end


#p people.size

# 统计结果
groups = [[],[],[],[],[],[],[],[]]
scores_of_group = Array.new(8,0)
people.each do |id,user|
  #p "#{user.dui} #{user.name}: #{user.scores}"
  arr = Array.new(7)
  arr[0] = id
  arr[1] = user.name
  arr[2] = user.scores
  arr[3] = user.scores_dabiao
  arr[4] = user.scores_dati
  arr[5] = user.dabiao
  arr[6] = user.dati
  #p arr
  #p user.dui
  case user.dui
  when "一队"
    groups[0].push  arr
    scores_of_group[0] += user.scores.to_i
  when "二队"
    groups[1].push arr
    scores_of_group[1] += user.scores.to_i
  when "三队"
    groups[2].push arr
    scores_of_group[2] += user.scores.to_i
  when "四队"
    groups[3].push arr
    scores_of_group[3] += user.scores.to_i
  when "五队"
    groups[4].push arr
    scores_of_group[4] += user.scores.to_i
  when "六队"
    groups[5].push arr
    scores_of_group[5] += user.scores.to_i
  when "七队"
    groups[6].push arr
    scores_of_group[6] += user.scores.to_i
  when "八队"
    groups[7].push arr
    scores_of_group[7] += user.scores.to_i
  end
end
#p scores_of_group
#p groups[0].size

#输出报表
puts "健步行比赛分数统计表"
puts ""

8.times do |i|
  puts "第#{i+1}队 总分 #{scores_of_group[i]}"
  puts "--详细数据："
  puts "-----ID----------姓名-----总分-----达标-----答题----- "
  groups[i].each do |user|
    #puts "-----ID-----姓名-----总分-----达标-----答题----- "
    puts "-----#{user[0]}-----#{user[1]}-----#{user[2]}-----#{user[3]}-----#{user[4]}"
    print "----------  达标日期:"
    user[5].each do |ii|
      print "#{ii}, "
    end
    puts ""
    print "---------- 答题得分："
    user[6].each {|ii| print "#{ii}, "}
    puts ""
  end
end
