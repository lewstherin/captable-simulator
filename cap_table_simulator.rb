require 'spreadsheet'
require 'yaml'

class CapTableSimulator
  attr_accessor :founders, :rounds, :option_pool, :cap_table, :history

  def initialize
    @founders = {}
    @rounds = []
    @option_pool = {}
    @cap_table = {}
    @history = []
  end

  def start
    menu
  end

  def ask_founders_info
    puts "Enter founder names and their equity distribution (comma separated, e.g., John:50, Jane:50):"
    input = gets.chomp
    input.split(',').each do |founder|
      name, equity = founder.split(':')
      @founders[name.strip] = equity.to_f
    end
    puts "Enter initial stock option pool percentage:"
    initial_pool = gets.chomp.to_f
    @option_pool['Initial Pool'] = initial_pool
    update_cap_table
  end

  def update_cap_table
    @cap_table = @founders.merge(@option_pool)
    @rounds.each do |round|
      @cap_table.merge!(round[:investors])
    end
    recalculate_percentages
  end

  def recalculate_percentages
    total_percentage = @cap_table.values.sum
    @cap_table.each do |holder, percentage|
      @cap_table[holder] = (percentage / total_percentage) * 100
    end
  end

  def menu
    loop do
      puts "\nCap Table Menu:"
      puts "1. Initialize cap table with founder information"
      puts "2. Add funding round (SAFE, convertible note, or priced equity round)"
      puts "3. Allocate common stock or option grants to employees"
      puts "4. Create stock option pool"
      puts "5. View cap table"
      puts "6. Download cap table as Excel"
      puts "7. Undo last round"
      puts "8. Save cap table to file"
      puts "9. Load cap table from file"
      puts "10. Exit"
      choice = gets.chomp.to_i

      case choice
      when 1
        ask_founders_info
      when 2
        add_funding_round
      when 3
        allocate_stock_or_options
      when 4
        create_stock_option_pool
      when 5
        view_cap_table
      when 6
        download_cap_table
      when 7
        undo_last_round
      when 8
        save_cap_table_to_file
      when 9
        load_cap_table_from_file
      when 10
        break
      else
        puts "Invalid choice. Please try again."
      end
    end
  end

  def add_funding_round
    puts "Enter round name:"
    round_name = gets.chomp
    puts "Enter type of round (SAFE, Convertible Note, Priced Equity):"
    round_type = gets.chomp.downcase

    case round_type
    when 'safe'
      add_safe_round(round_name)
    when 'convertible note'
      add_convertible_note_round(round_name)
    when 'priced equity'
      add_priced_equity_round(round_name)
    else
      puts "Invalid round type. Please try again."
    end
  end

  def add_safe_round(round_name)
    save_current_state
    puts "Enter amount invested:"
    amount_invested = gets.chomp.to_f
    puts "Enter discount rate (as a percentage, leave blank if MFN SAFE):"
    discount_rate_input = gets.chomp
    discount_rate = discount_rate_input.empty? ? nil : discount_rate_input.to_f
    puts "Enter post-money valuation cap (leave blank if not applicable):"
    post_money_cap_input = gets.chomp
    post_money_cap = post_money_cap_input.empty? ? nil : post_money_cap_input.to_f

    if discount_rate.nil? && post_money_cap.nil?
      puts "This is a MFN SAFE. The shares will be allocated in the next equity round."
      @rounds << { name: round_name, type: 'MFN SAFE', invested: amount_invested, investors: { round_name => 0 }, new_shares_percentage: 0 }
    else
      if post_money_cap
        post_money_valuation = post_money_cap
      else
        post_money_valuation = @cap_table.values.sum * (1 - discount_rate / 100)
      end

      new_shares_percentage = (amount_invested / post_money_valuation) * 100
      @rounds << { name: round_name, type: 'SAFE', invested: amount_invested, discount_rate: discount_rate, post_money_cap: post_money_cap, new_shares_percentage: new_shares_percentage, investors: { round_name => new_shares_percentage } }
    end
    update_cap_table
  end

  def add_convertible_note_round(round_name)
    save_current_state
    puts "Enter amount invested:"
    amount_invested = gets.chomp.to_f
    puts "Enter conversion price:"
    conversion_price = gets.chomp.to_f
    @rounds << { name: round_name, type: 'Convertible Note', invested: amount_invested, conversion_price: conversion_price, investors: { round_name => amount_invested } }
    update_cap_table
  end

  def add_priced_equity_round(round_name)
    save_current_state
    puts "Enter pre-money valuation:"
    pre_money_valuation = gets.chomp.to_f
    puts "Enter amount invested:"
    amount_invested = gets.chomp.to_f
    post_money_valuation = pre_money_valuation + amount_invested
    new_shares_percentage = (amount_invested / post_money_valuation) * 100

    puts "Enter new stock option pool percentage after the round:"
    new_pool_percentage = gets.chomp.to_f
    create_stock_option_pool_pre_money(new_pool_percentage)

    @rounds << { name: round_name, type: 'Priced Equity', pre_money: pre_money_valuation, invested: amount_invested, post_money: post_money_valuation, new_shares_percentage: new_shares_percentage, investors: { round_name => new_shares_percentage } }
    handle_mfn_safe_conversion(post_money_valuation)
    update_cap_table
  end

  def handle_mfn_safe_conversion(post_money_valuation)
    @rounds.each do |round|
      if round[:type] == 'MFN SAFE'
        round_invested = round[:invested]
        mfn_new_shares_percentage = (round_invested / post_money_valuation) * 100
        round[:new_shares_percentage] = mfn_new_shares_percentage
        round[:investors][round[:name]] = mfn_new_shares_percentage # Updated line to correctly set the investor percentage
      end
    end
  end

  def create_stock_option_pool_pre_money(new_pool_percentage)
    existing_shareholders_percentage = 100.0 - new_pool_percentage
    @cap_table.each do |holder, percentage|
      @cap_table[holder] = (percentage / existing_shareholders_percentage) * 100
    end
    @option_pool["Pool After Round #{@rounds.size + 1}"] = new_pool_percentage
    update_cap_table
  end

  def allocate_stock_or_options
    save_current_state
    puts "Enter the name of the employee or entity:"
    name = gets.chomp
    puts "Enter the percentage of stock or options to allocate:"
    percentage = gets.chomp.to_f
    @option_pool[name] = percentage
    update_cap_table
  end

  def create_stock_option_pool
    save_current_state
    puts "Enter new stock option pool percentage:"
    new_pool_percentage = gets.chomp.to_f
    total_percentage = 100 - new_pool_percentage
    @cap_table.each do |holder, percentage|
      @cap_table[holder] = (percentage / total_percentage) * 100
    end
    @option_pool["Pool After Round #{@rounds.size + 1}"] = new_pool_percentage
    update_cap_table
  end

  def undo_last_round
    if @history.empty?
      puts "No rounds to undo."
    else
      @rounds = @history.pop
      update_cap_table
      puts "Last round has been undone."
    end
  end

  def save_current_state
    @history << Marshal.load(Marshal.dump(@rounds))
  end

  def view_cap_table
    puts "\nCurrent Cap Table:"
    @cap_table.each do |holder, percentage|
      puts "#{holder}: #{percentage.round(2)}%"
    end
  end

  def download_cap_table
    book = Spreadsheet::Workbook.new
    sheet = book.create_worksheet
    sheet.row(0).concat %w[Holder Percentage]
    @cap_table.each_with_index do |(holder, percentage), index|
      sheet.row(index + 1).push holder, percentage
    end
    book.write 'cap_table.xls'
    puts "Cap table downloaded as cap_table.xls"
  end

  def save_cap_table_to_file
    puts "Enter the filename to save the cap table:"
    filename = gets.chomp
    File.open(filename, 'w') { |file| file.write(YAML.dump(self)) }
    puts "Cap table saved to #{filename}."
  end

  def load_cap_table_from_file
    puts "Enter the filename to load the cap table:"
    filename = gets.chomp
    if File.exist?(filename)
      loaded_data = YAML.safe_load(File.read(filename), permitted_classes: [CapTableSimulator, Symbol])
      @founders = loaded_data.founders
      @rounds = loaded_data.rounds
      @option_pool = loaded_data.option_pool
      @cap_table = loaded_data.cap_table
      @history = loaded_data.history
      update_cap_table
      puts "Cap table loaded from #{filename}."
    else
      puts "File not found."
    end
  end
end

# Start the cap table simulation
CapTableSimulator.new.start
