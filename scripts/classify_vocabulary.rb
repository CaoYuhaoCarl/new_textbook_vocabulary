#!/usr/bin/env ruby

require "csv"
require "fileutils"

INPUT_PATH = File.expand_path("../merged_vocabulary.csv", __dir__)
EXAM_PATH = File.expand_path("../zk_vocabulary_v2024_cleaned.md", __dir__)
OUTPUT_DIR = File.expand_path("../outputs/019f8ca8-fe4a-7000-8d6b-d1ed952617fa", __dir__)

CLASS_ACTIONS = {
  "A核心" => "保留在主词表",
  "B认读" => "保留在认读区",
  "C剔除" => "从学生主词表剔除",
  "D复核" => "等待人工确认"
}.freeze

COUNTRIES_AND_CONTINENTS = [
  "Africa",
  "America",
  "Antarctica",
  "Argentina",
  "Asia",
  "Australia",
  "Belgium",
  "Brazil",
  "Canada",
  "Central America",
  "Egypt",
  "Equatorial Guinea",
  "Europe",
  "Finland",
  "France",
  "Germany",
  "Greece",
  "India",
  "Indonesia",
  "Ireland",
  "Italy",
  "Japan",
  "Kenya",
  "Malaysia",
  "Mexico",
  "Netherlands",
  "Peru",
  "Philippines",
  "Poland",
  "Russia",
  "Scotland",
  "Singapore",
  "Thailand",
  "Türkiye",
  "UK",
  "US",
  "USA"
].freeze

COMMON_FESTIVALS_TO_REVIEW = [
  "April Fool's Day",
  "Dragon Boat Festival",
  "Mid-Autumn Festival",
  "National Day",
  "Tomb-sweeping Day"
].freeze

LOW_VALUE_CULTURAL_TERMS = [
  "Auld Lang Syne",
  "Beijing roast duck",
  "CPC Founding Day",
  "Dongpo pork",
  "Gongbao chicken",
  "Guest-Greeting / Pine",
  "Guoqiao Rice Noodles",
  "Hauptschule",
  "Holi",
  "Mapo tofu",
  "PLA Day",
  "Realschule",
  "Sea of Clouds",
  "Sham El-Nessim"
].freeze

CORE_EXPRESSIONS = [
  "Congratulations (on ...)!",
  "How about ...?",
  "What about …?",
  "You're welcome.",
  "Chinese chess",
  "play Chinese chess"
].freeze

RECOGNITION_TERMS = [
  "AI",
  "AM",
  "BCE",
  "CE",
  "Chinese lunar calendar",
  "CT scan",
  "DNA test",
  "Dr",
  "HSK",
  "IT",
  "Nobel Prize",
  "PM",
  "PRC",
  "PLA",
  "Pulitzer Prize",
  "UN",
  "VR",
  "WHO",
  "World War II"
].freeze

ORGANIZATIONS = [
  "Blue Sky Rescue",
  "United Nations Peacekeeping"
].freeze

WORK_TITLES = [
  "A Dream of Red Mansions",
  "Alice's Adventures in Wonderland",
  "Around the World in Eighty Days",
  "Harry Potter and the Philosopher's Stone",
  "Journey to the West",
  "Mona Lisa",
  "My Childhood",
  "Oliver Twist",
  "Outlaws of the Marsh",
  "Red Star over China",
  "The Adventures of Tom Sawyer",
  "The Final Problem",
  "The Little Match Girl",
  "The Little Prince",
  "The Old Man and the Sea",
  "The Romance of the Three Kingdoms",
  "The Secret Garden",
  "The Three-Body Problem",
  "The Time Machine",
  "The Wonderful Wizard of Oz",
  "Treasure Island",
  "Venus de Milo"
].freeze

PERSON_NAMES = [
  "Adam", "Alan Turing", "Alexander Bell", "Alice", "Allen", "Amelia Earhart",
  "Ann", "Anna", "Anne", "Arthur Fry", "Asimov", "Badal", "Baker", "Ben",
  "Beth", "Betty", "Bill", "Billy", "Bing Dwen Dwen", "Bob", "Brown", "Bruno",
  "Cathy", "Charles Dickens", "Clarice", "Clark", "Coco", "Colin", "Confucius",
  "Cooper", "Craven", "Da Vinci", "Dave", "David", "Davis", "Debbie", "Diana",
  "Dickon", "Edward", "Ella", "Emma", "Eric", "Eric Moussambani", "Ernest Hemingway",
  "Flora", "Florence Nightingale", "Frank", "Frances Hodgson Burnett", "Fred",
  "Frédéric Chopin", "George Mallory", "Green", "Guglielmo Marconi", "Halla",
  "Hans Christian Andersen", "Harry", "Helen", "I.M. Pei", "Irène", "Isaac Newton",
  "Isambard Brunel", "Jack", "James", "Jane", "Jason", "Jeff", "Jennifer", "Jenny",
  "Ji-Hoon", "Jim", "Joe", "John", "Johnny", "Jones", "Judy", "Julie", "Kaito",
  "Kate", "Kelly", "King Tut", "Kukulkan", "Lee", "Lennox", "Lewis Carroll", "Lily",
  "Linda", "Lisa", "Luca", "Lucy", "Malee", "Mandy", "Marie", "Marie Curie",
  "Mark Twain", "Mary", "Matt", "Maya", "Michael Faraday", "Mike", "Miller", "Nancy",
  "Orville Wright", "Oscar", "Pauline", "Peter", "Rick", "Robert", "Robert Schumann",
  "Rose", "Sally", "Sam", "Sarah", "Sherlock Holmes", "Smith", "Spencer Silver",
  "Stephen", "Steve", "Susan", "Terry", "Thompson", "Tilly", "Timo", "Tina", "Tom",
  "Tony", "Usain Bolt", "Victor", "Victor Hugo", "Vincent", "Vincent van Gogh",
  "Wilbur Wright", "William Shakespeare", "Wood"
].freeze

SPECIFIC_PLACES = [
  "Alexander Garden", "Angel Falls", "Astor Garden Court", "Atlantic Ocean", "Big Ben",
  "Bright Peak", "Cairo", "Central Park", "Dead Sea", "East African Rift Valley",
  "Great Barrier Reef", "Great Pyramid", "Helsinki", "Hong Kong", "Inga Falls",
  "Lake Baikal", "Lhasa", "London", "Louvre Museum", "Mariana Trench", "Mars",
  "Metropolitan Museum of Art", "Mogao Caves", "Mombasa", "Moscow", "Moscow Metro",
  "Mount Huangshan", "Mount Kilimanjaro", "Mount Qomolangma", "Nairobi", "New York",
  "Nile River", "Oakland", "Orchid Pavilion", "Paris", "Phuket", "Proxima Centauri",
  "Red Square", "Rio de Janeiro", "Sahara Desert", "Seattle", "Siberia", "Sphinx",
  "Stockholm", "Stonehenge", "Sydney", "Taklimakan Desert", "Tanggula Pass",
  "Terracotta Army", "The Arctic Ocean", "Times Square", "Victoria Falls", "Vienna",
  "Vancouver", "Voyager 1", "Western Regions", "Yangtze River", "the Great Wall", "the Silk Road",
  "the Victory Museum", "the Warring States Period"
].freeze

def normalize(value)
  value.to_s.downcase.strip.gsub(/[[:space:]\u00a0]+/, " ").gsub("…", "...")
end

def load_exam_words
  exam_words = {}
  File.foreach(EXAM_PATH, encoding: "utf-8") do |line|
    headword_match = line.match(/^\*\*([^*]+)\*\*/)
    next unless headword_match

    headword = headword_match[1].strip
    frequency_match = line.match(/\[词频:\s*(\d+)\]/)
    frequency = frequency_match && frequency_match[1].to_i
    key = normalize(headword)
    if !exam_words.key?(key) || (frequency && (!exam_words[key] || frequency < exam_words[key]))
      exam_words[key] = frequency
    end
  end
  exam_words
end

def capitalized_candidate?(word)
  word.match?(/[A-Z]/)
end

def multiword?(word)
  word.match?(/\s/)
end

def classify(row, exam_words)
  word = row.fetch("word").strip
  part_of_speech = row.fetch("part_of_speech", "").to_s.strip
  definition = row.fetch("chinese_definition", "").to_s.strip
  exam_key = normalize(word)
  exam_frequency = exam_words[exam_key]
  exam_match = exam_words.key?(exam_key)

  if CORE_EXPRESSIONS.include?(word)
    return ["A核心", "核心表达", "教材明确列出的常用交际表达或固定搭配", ""]
  end

  if WORK_TITLES.include?(word) || definition.start_with?("《")
    return ["C剔除", "作品名称", "作品或艺术品专名，不作为学生背诵词条", "可保留在文化附录"]
  end

  if PERSON_NAMES.include?(word) || (capitalized_candidate?(word) && definition.match?(/人名|姓氏/))
    note = exam_match ? "与考频表普通词同形，按本条中文释义判定为人名" : ""
    return ["C剔除", "人名", "人物名、姓氏或虚构人物名", note]
  end

  if COUNTRIES_AND_CONTINENTS.include?(word)
    return ["D复核", "国家或大洲", "属于地名，但常见国家和大洲名称可能具有认读价值", "确认是剔除还是放入认读区"]
  end

  if SPECIFIC_PLACES.include?(word) || word.match?(/(?:Mount|River|Falls|Sea|Desert|Trench|Lake|Ocean|Square|Park|Museum|Garden|Reef|Valley|Pass|Caves)\b/)
    return ["C剔除", "具体地名", "城市、景点、地理实体或历史地点专名", ""]
  end

  if ORGANIZATIONS.include?(word)
    return ["C剔除", "组织名称", "具体组织或机构名称，不作为学生背诵词条", "可保留在文化附录"]
  end

  if word == "Titanic"
    return ["C剔除", "命名实体", "具体船名，不作为学生背诵词条", ""]
  end

  if COMMON_FESTIVALS_TO_REVIEW.include?(word)
    return ["D复核", "常见节日", "属于专名，但可能在中考语境中出现", "确认是认读还是剔除"]
  end

  if LOW_VALUE_CULTURAL_TERMS.include?(word)
    return ["C剔除", "低考试价值文化词", "文化背景专名或外语教育制度名称，预计不要求拼写", "可保留在文化附录"]
  end

  if RECOGNITION_TERMS.include?(word)
    return ["B认读", "缩写或背景概念", "有课文理解价值，但不建议作为重点默写词", ""]
  end

  if word == "the Olympics"
    return ["D复核", "赛事专名", "常见赛事名称可能具有认读价值", "确认是认读还是剔除"]
  end

  if word == "Gymnasium"
    return ["C剔除", "低考试价值文化词", "本条为德国学校类型，不是体育馆这一普通词义", "避免与普通词义混淆"]
  end

  if capitalized_candidate?(word) && part_of_speech.empty?
    note = exam_match ? "考频表存在同形词，需以本条专名释义为准" : ""
    return ["C剔除", "其他专名", "大写专名且无普通词词性，预计不作为考试拼写词", note]
  end

  if capitalized_candidate?(word) && !part_of_speech.empty?
    if exam_match
      note = %w[TRUE FALSE].include?(word) ? "出版时建议恢复为教材常规小写形式" : ""
      return ["A核心", "普通词", "具有明确词性，并与浙江中考考频表词头匹配", note]
    end
    return ["B认读", "普通词", "具有明确词性，但未与浙江中考考频表词头直接匹配", ""]
  end

  if multiword?(word)
    if word == "starter unit"
      return ["B认读", "教材结构词", "用于教材结构说明，不作为重点短语", ""]
    end
    return ["A核心", "短语或固定表达", "教材明确列出的短语或句型，应独立掌握", "后续需结合短语考查价值再细分"]
  end

  if part_of_speech.empty?
    return ["B认读", "外语或文化用语", "无普通英语词性，主要用于课文理解", ""]
  end

  if exam_match
    return ["A核心", "普通词", "与浙江中考考频表词头直接匹配", ""]
  end

  ["B认读", "普通词", "教材词汇，但未与浙江中考考频表词头直接匹配", "需在样章阶段复核是否升级为核心"]
end

def write_csv(path, headers, rows)
  File.open(path, "wb") do |file|
    file.write("\uFEFF".encode("utf-8"))
    csv = CSV.new(file, write_headers: true, headers: headers)
    rows.each { |row| csv << headers.map { |header| row[header] } }
  end
end

exam_words = load_exam_words
source_rows = CSV.read(INPUT_PATH, headers: true, encoding: "bom|utf-8")

classified_rows = source_rows.each_with_index.map do |row, index|
  classification, entry_type, reason, review_note = classify(row, exam_words)
  word = row.fetch("word").strip
  exam_key = normalize(word)
  exam_frequency = exam_words[exam_key]
  exam_match = exam_words.key?(exam_key)

  {
    "entry_id" => format("NV%04d", index + 1),
    "source_row" => index + 2,
    "book" => row["book"],
    "unit" => row["unit"],
    "word" => word,
    "normalized_word" => normalize(word),
    "pronunciation" => row["pronunciation"],
    "part_of_speech" => row["part_of_speech"],
    "chinese_definition" => row["chinese_definition"],
    "page_number" => row["page_number"],
    "entry_form" => multiword?(word) ? "多词表达" : "单词",
    "capitalized_candidate" => capitalized_candidate?(word) ? "是" : "否",
    "exam_headword_match" => exam_match ? "是" : "否",
    "exam_frequency_rank" => exam_frequency,
    "classification" => classification,
    "action" => CLASS_ACTIONS.fetch(classification),
    "entry_type" => entry_type,
    "decision_reason" => reason,
    "review_note" => review_note,
    "source_file" => row["source_file"]
  }
end

headers = classified_rows.first.keys
FileUtils.mkdir_p(OUTPUT_DIR)
write_csv(File.join(OUTPUT_DIR, "vocabulary_classification.csv"), headers, classified_rows)
write_csv(
  File.join(OUTPUT_DIR, "suggested_exclusions.csv"),
  headers,
  classified_rows.select { |row| row["classification"] == "C剔除" }
)
write_csv(
  File.join(OUTPUT_DIR, "manual_review.csv"),
  headers,
  classified_rows.select { |row| row["classification"] == "D复核" }
)

final_rows = classified_rows.map do |row|
  final_row = row.dup
  if final_row["classification"] == "D复核"
    final_row["classification"] = "B认读"
    final_row["action"] = CLASS_ACTIONS.fetch("B认读")
    final_row["decision_reason"] = "经确认保留为认读词，不要求重点默写"
    final_row["review_note"] = "原D类已统一转入B类认读区"
    final_row["final_decision_note"] = "用户确认：常见国家、大洲、节日和赛事保留为认读词"
  else
    final_row["final_decision_note"] = "沿用首轮分类"
  end
  final_row
end

final_headers = final_rows.first.keys
write_csv(
  File.join(OUTPUT_DIR, "vocabulary_classification_final.csv"),
  final_headers,
  final_rows
)
write_csv(
  File.join(OUTPUT_DIR, "vocabulary_book_source.csv"),
  final_headers,
  final_rows.reject { |row| row["classification"] == "C剔除" }
)
write_csv(
  File.join(OUTPUT_DIR, "final_exclusions.csv"),
  final_headers,
  final_rows.select { |row| row["classification"] == "C剔除" }
)

summary_headers = ["book", "classification", "count"]
summary_rows = []
["全部", "7A", "7B", "8A", "8B", "9A", "9B"].each do |book|
  rows_for_book = book == "全部" ? classified_rows : classified_rows.select { |row| row["book"] == book }
  CLASS_ACTIONS.keys.each do |classification|
    summary_rows << {
      "book" => book,
      "classification" => classification,
      "count" => rows_for_book.count { |row| row["classification"] == classification }
    }
  end
end
write_csv(File.join(OUTPUT_DIR, "classification_summary.csv"), summary_headers, summary_rows)

final_summary_rows = []
["全部", "7A", "7B", "8A", "8B", "9A", "9B"].each do |book|
  rows_for_book = book == "全部" ? final_rows : final_rows.select { |row| row["book"] == book }
  CLASS_ACTIONS.keys.each do |classification|
    final_summary_rows << {
      "book" => book,
      "classification" => classification,
      "count" => rows_for_book.count { |row| row["classification"] == classification }
    }
  end
end
write_csv(File.join(OUTPUT_DIR, "final_classification_summary.csv"), summary_headers, final_summary_rows)

raise "Unexpected source row count" unless source_rows.length == 2_681
raise "Classification lost rows" unless classified_rows.length == source_rows.length
raise "Final classification lost rows" unless final_rows.length == source_rows.length
raise "Unexpected final retained count" unless final_rows.count { |row| row["classification"] != "C剔除" } == 2_445
raise "Final classification still contains review rows" if final_rows.any? { |row| row["classification"] == "D复核" }
raise "Missing decision reason" if classified_rows.any? { |row| row["decision_reason"].to_s.empty? }
raise "Unknown classification" if classified_rows.any? { |row| !CLASS_ACTIONS.key?(row["classification"]) }

expected_cases = {
  "Helen" => ["C剔除", "人名"],
  "London" => ["C剔除", "具体地名"],
  "Australia" => ["D复核", "国家或大洲"],
  "You're welcome." => ["A核心", "核心表达"],
  "Brown" => ["C剔除", "人名"],
  "AI" => ["B认读", "缩写或背景概念"],
  "Asia" => ["D复核", "国家或大洲"],
  "Gymnasium" => ["C剔除", "低考试价值文化词"]
}

expected_cases.each do |word, expected|
  found = classified_rows.find { |row| row["word"] == word }
  raise "Missing verification case: #{word}" unless found
  actual = [found["classification"], found["entry_type"]]
  raise "Unexpected classification for #{word}: #{actual.inspect}" unless actual == expected
end

common_rose = classified_rows.find do |row|
  row["word"] == "rose" && row["part_of_speech"].to_s.include?("n.")
end
raise "Common noun rose should remain in the book" unless common_rose && common_rose["classification"] != "C剔除"

counts = CLASS_ACTIONS.keys.to_h do |classification|
  [classification, classified_rows.count { |row| row["classification"] == classification }]
end

puts "source_rows=#{source_rows.length}"
puts counts.map { |classification, count| "#{classification}=#{count}" }.join(" ")
puts "exam_headword_matches=#{classified_rows.count { |row| row["exam_headword_match"] == "是" }}"
puts "output_dir=#{OUTPUT_DIR}"
