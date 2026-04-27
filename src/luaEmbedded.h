// This file is part of the MultiReplace plugin for Notepad++.
// Copyright (C) 2023 Thomas Knoefel
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program. If not, see <https://www.gnu.org/licenses/>.

#ifndef LUA_EMBEDDED_H
#define LUA_EMBEDDED_H


// ---------- Original Lua Source Code  ----------
static const char* luaSourceCode = R"(
------------------------------------------------------------------
-- 1) cond function
------------------------------------------------------------------
function cond(condition, trueVal, falseVal)
  if condition == nil then
    error('cond: condition cannot be nil')
  end
  if condition and trueVal == nil then
    error('cond: trueVal cannot be nil')
  end

  -- Resolve thunks: if trueVal/falseVal is a function, call it for the value.
  if type(trueVal) == 'function' then
    trueVal = trueVal()
  end
  if type(falseVal) == 'function' then
    falseVal = falseVal()
  end

  -- Pick the active branch. Cannot use "a and b or c" here: it would
  -- collapse legitimate falsy picks (false, '', 0 are fine values too).
  local picked
  if condition then
    picked = trueVal
  else
    picked = falseVal
  end

  local res
  if type(picked) == 'table' then
    -- Forward { result, skip } from the chosen branch as-is.
    res = { result = picked.result, skip = picked.skip }
  elseif picked == nil then
    -- Falsy condition without a falseVal: skip the replacement entirely.
    res = { result = '', skip = true }
  else
    res = { result = picked, skip = false }
  end

  resultTable = res
  return res
end

------------------------------------------------------------------
-- 2) set function
------------------------------------------------------------------
function set(strOrCalc)
  if strOrCalc == nil then
    error('set: cannot be nil')
  end

  local res
  if type(strOrCalc) == 'string' then
    res = { result = strOrCalc, skip = false }
  elseif type(strOrCalc) == 'number' then
    res = { result = tostring(strOrCalc), skip = false }
  else
    error('set: Expected string or number')
  end

  resultTable = res
  return res
end

------------------------------------------------------------------
-- 3) fmtN function (helper, intended to be wrapped in set()/cond())
-- Returns a formatted number STRING. Does NOT touch resultTable.
------------------------------------------------------------------
function fmtN(num, maxDecimals, fixedDecimals)
  if type(num) ~= 'number' then
    error('fmtN: num must be a number')
  end
  if type(maxDecimals) ~= 'number' then
    error('fmtN: maxDecimals must be a number')
  end
  if type(fixedDecimals) ~= 'boolean' then
    error('fmtN: fixedDecimals must be a boolean')
  end

  local multiplier = 10 ^ maxDecimals
  local rounded = math.floor(num * multiplier + 0.5) / multiplier

  if fixedDecimals then
    return string.format('%.' .. maxDecimals .. 'f', rounded)
  end

  -- Variable-decimals mode: drop the fractional part if it rounds away.
  local intPart, fracPart = math.modf(rounded)
  if fracPart == 0 then
    return tostring(intPart)
  end
  return tostring(rounded)
end

------------------------------------------------------------------
-- 4) vars function (and init alias)
-- Sets global variables with first-write-wins semantics: re-running
-- the same Lua snippet won't clobber values from a previous run.
------------------------------------------------------------------
function vars(args)
  for name, value in pairs(args) do
    -- Only set the global if it doesn't already exist; this lets a
    -- caller re-run the snippet without resetting earlier state.
    if _G[name] == nil then
      if type(name) ~= 'string' then
        error('vars: Variable name must be a string')
      end
      if not string.match(name, '^[A-Za-z_][A-Za-z0-9_]*$') then
        error('vars: Invalid variable name \"' .. tostring(name) .. '\"')
      end
      if value == nil then
        error('vars: Value missing for variable \"' .. tostring(name) .. '\"')
      end
      -- In regex replacement mode, escape backslashes so user values
      -- substituted into the regex stay literal.
      if type(value) == 'string' and REGEX then
        value = value:gsub('\\\\', '\\\\\\\\')
      end
      _G[name] = value
    end
  end

  -- Initialize resultTable on first call; preserve any values already
  -- placed there by an earlier function call in the same snippet.
  resultTable = resultTable or { result = '', skip = true }
  if resultTable.result == nil then resultTable.result = '' end
  if resultTable.skip == nil then resultTable.skip = true end
  return resultTable
end

init = vars  -- 'init' alias for compatibility

------------------------------------------------------------------
-- 5) lkp function
------------------------------------------------------------------
hashTables = {}

function lkp(key, hpath, inner)
  local res = { result = '', skip = false }

  if key == nil then
    error('lkp: key passed to file is nil in ' .. tostring(hpath))
  end
  if type(key) == 'number' then
    key = tostring(key)
  end

  if hpath == nil or hpath == '' then
    error('lkp: file path is invalid or empty')
  end
  if inner == nil then
    inner = false
  end

  -- Lazy-load + cache the file's lookup table on first use.
  local tbl = hashTables[hpath]
  if tbl == nil then
    local success, dataEntries = safeLoadFileSandbox(hpath)
    if not success then
      error('lkp: failed to safely load file at ' .. tostring(hpath) .. ': ' .. tostring(dataEntries))
    end
    if type(dataEntries) ~= 'table' then
      error('lkp: invalid format in file at ' .. tostring(hpath))
    end

    tbl = {}
    for _, entry in ipairs(dataEntries) do
      local keys = entry[1]
      local value = entry[2]

      if value ~= nil then
        local kt = type(keys)
        if kt == 'table' then
          for _, k in ipairs(keys) do
            if type(k) == 'number' then
              k = tostring(k)
            end
            tbl[k] = value
          end
        elseif kt == 'string' then
          tbl[keys] = value
        elseif kt == 'number' then
          tbl[tostring(keys)] = value
        end
      end
    end
    hashTables[hpath] = tbl
  end

  -- Hot path: hash lookup.
  local val = tbl[key]
  if val == nil then
    if inner then
      res.result = nil
    else
      res.result = key
    end
  else
    res.result = val
  end

  resultTable = res
  return res
end

------------------------------------------------------------------
-- 6) lvars function
-- File-based variant of vars(): loads a table of name=value pairs
-- from a .lua file via the sandbox loader.
------------------------------------------------------------------
function lvars(filePath)
  if filePath == nil or filePath == '' then
    error('lvars: file path is invalid or empty')
  end

  local success, dataTable = safeLoadFileSandbox(filePath)
  if not success then
    error('lvars: failed to safely load file at ' .. tostring(filePath) .. ': ' .. tostring(dataTable))
  end
  if type(dataTable) ~= 'table' then
    error('lvars: invalid data format in file at ' .. tostring(filePath))
  end

  for name, value in pairs(dataTable) do
    if type(name) ~= 'string' then
      error('lvars: Variable name must be a string. Found invalid key \"' .. tostring(name) .. '\"')
    end
    if not string.match(name, '^[A-Za-z_][A-Za-z0-9_]*$') then
      error('lvars: Invalid variable name \"' .. tostring(name) .. '\"')
    end
    if value == nil then
      error('lvars: Value missing for variable \"' .. tostring(name) .. '\"')
    end
    if REGEX and type(value) == 'string' then
      value = value:gsub('\\\\', '\\\\\\\\')
    end
    _G[name] = value
  end

  resultTable = resultTable or { result = '', skip = true }
  if resultTable.result == nil then resultTable.result = '' end
  if resultTable.skip == nil then resultTable.skip = true end
  return resultTable
end

------------------------------------------------------------------
-- 7) Helper Functions: trim, padL, padR, tonum
------------------------------------------------------------------

-- Removes leading and trailing whitespace
function trim(s)
  if s == nil then return "" end
  return (tostring(s):gsub("^%s+", ""):gsub("%s+$", ""))
end

-- Pads string 's' on the LEFT with char 'c' until width 'w' is reached
-- Usage: padL(CNT, 3, "0") -> "001"
function padL(s, w, c)
  s = tostring(s or "")
  w = tonumber(w) or 0
  c = tostring(c or " ")
  if #c == 0 then c = " " end
  
  if #s >= w then 
    return s 
  end
  
  return string.rep(c, w - #s) .. s
end

-- Pads string 's' on the RIGHT with char 'c' until width 'w' is reached
-- Usage: padR(match, 20, " ")
function padR(s, w, c)
  s = tostring(s or "")
  w = tonumber(w) or 0
  c = tostring(c or " ")
  if #c == 0 then c = " " end
  
  if #s >= w then 
    return s 
  end
  
  return s .. string.rep(c, w - #s)
end

-- Converts string to number, accepting both dot (.) and comma (,) as decimal separator
-- Usage: tonum(CAP1) * 2, tonum("3,14") -> 3.14
function tonum(s)
  if s == nil then
    return nil
  end
  if type(s) == 'number' then
    return s
  end
  if type(s) ~= 'string' then
    s = tostring(s)
  end
  -- Replace comma with dot for decimal separator compatibility
  s = s:gsub(',', '.')
  return tonumber(s)
end

------------------------------------------------------------------
-- 8) lcmd function
-- Loads a file that returns a table of helper functions and registers
-- them globally. Idempotent: subsequent calls with the same path are
-- no-ops, so re-running the same Lua snippet won't trip the
-- "command already exists" guard.
------------------------------------------------------------------
loadedCmdFiles = {}

function lcmd(path)
  if loadedCmdFiles[path] then
    resultTable = resultTable or { result = "", skip = true }
    return resultTable
  end

  local ok, mod = safeLoadFileSandbox(path)
  if not ok then
    error("lcmd: " .. tostring(mod))
  end
  if type(mod) ~= "table" then
    error("lcmd: file must return a table of functions")
  end

  local env = _ENV or _G
  local count = 0
  for name, fn in pairs(mod) do
    if type(fn) == "function" then
      if env[name] ~= nil then
        error("lcmd: command already exists: " .. tostring(name))
      end
      env[name] = fn
      count = count + 1
    end
  end
  if count == 0 then
    error("lcmd: file exported no functions")
  end

  loadedCmdFiles[path] = true

  resultTable = resultTable or { result = "", skip = true }
  return resultTable
end

)";
static const size_t luaSourceSize = sizeof(luaSourceCode);

#endif // LUA_EMBEDDED_H