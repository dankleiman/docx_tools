require 'simplecov'
SimpleCov.start

require 'bundler/setup'

$LOAD_PATH.unshift File.expand_path('../../lib', __FILE__)
require 'docx_tools'

require 'minitest/autorun'
