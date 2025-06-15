#ifndef LUA_EMBEDDED_H
#define LUA_EMBEDDED_H


// ---------- Original Lua Source Code  ----------
static const char* luaSourceCode = R"(
------------------------------------------------------------------
-- 1) cond function
------------------------------------------------------------------
function cond(cond, trueVal, falseVal)
  local res = {result = '', skip = false}  -- Initialize result table with defaults
  if cond == nil then
    error('cond: condition cannot be nil')
    return res
  end

  if cond and trueVal == nil then
    error('cond: trueVal cannot be nil')
    return res
  end

  if not cond and falseVal == nil then
    res.skip = true  -- Skip ONLY if falseVal is not provided
  end

  if type(trueVal) == 'function' then
    trueVal = trueVal()
  end

  if type(falseVal) == 'function' then
    falseVal = falseVal()
  end

  if cond then
    if type(trueVal) == 'table' then
      res.result = trueVal.result
      res.skip = trueVal.skip
    else
      res.result = trueVal
      res.skip = false
    end
  else
    if not res.skip then
      if type(falseVal) == 'table' then
        res.result = falseVal.result
        res.skip = falseVal.skip
      else
        res.result = falseVal
      end
    end
  end
  resultTable = res
  return res
end

------------------------------------------------------------------
-- 2) set function
------------------------------------------------------------------
function set(strOrCalc)
  local res = {result = '', skip = false}
  if strOrCalc == nil then
    error('set: cannot be nil')
    return
  end
  if type(strOrCalc) == 'string' then
    res.result = strOrCalc
  elseif type(strOrCalc) == 'number' then
    res.result = tostring(strOrCalc)
  else
    error('set: Expected string or number')
    return
  end
  resultTable = res
  return res
end

------------------------------------------------------------------
-- 3) fmtN function
------------------------------------------------------------------
function fmtN(num, maxDecimals, fixedDecimals)
  if num == nil then
    error('fmtN: num cannot be nil')
    return
  elseif type(num) ~= 'number' then
    error('fmtN: Invalid type for num. Expected a number')
    return
  end
  if maxDecimals == nil then
    error('fmtN: maxDecimals cannot be nil')
    return
  elseif type(maxDecimals) ~= 'number' then
    error('fmtN: Invalid type for maxDecimals. Expected a number')
    return
  end
  if fixedDecimals == nil then
    error('fmtN: fixedDecimals cannot be nil')
    return
  elseif type(fixedDecimals) ~= 'boolean' then
    error('fmtN: Invalid type for fixedDecimals. Expected a boolean')
    return
  end

  local multiplier = 10 ^ maxDecimals
  local rounded = math.floor(num * multiplier + 0.5) / multiplier
  local output = ''
  if fixedDecimals then
    output = string.format('%.' .. maxDecimals .. 'f', rounded)
  else
    local intPart, fracPart = math.modf(rounded)
    if fracPart == 0 then
      output = tostring(intPart)
    else
      output = tostring(rounded)
    end
  end
  return output
end

------------------------------------------------------------------
-- 4) vars function (and init alias)
------------------------------------------------------------------
function vars(args)
  for name, value in pairs(args) do
    -- Set the global variable only if it does not already exist
    if _G[name] == nil then
      if type(name) ~= 'string' then
        error('vars: Variable name must be a string')
      end
      if not string.match(name, '^[A-Za-z_][A-Za-z0-9_]*$') then
        error('vars: Invalid variable name')
      end
      if value == nil then
        error('vars: Value missing')
      end
      -- If REGEX is true and value is a string, escape backslashes
      if type(value) == 'string' and REGEX then
        value = value:gsub('\\\\', '\\\\\\\\')
      end
      _G[name] = value
    end
  end

  -- Forward or initialize resultTable
  local res = {result = '', skip = true}
  if resultTable == nil then
    resultTable = res
  else
    if resultTable.result == nil then
      resultTable.result = res.result
    end
    if resultTable.skip == nil then
      resultTable.skip = res.skip
    end
  end
  return resultTable
end

init = vars  -- 'init' alias for compatibility

------------------------------------------------------------------
-- 5) lkp function
------------------------------------------------------------------
hashTables = {}

function lkp(key, hpath, inner)
  local res = { result = '', skip = false }

  if type(key) == 'number' then
    key = tostring(key)
  end

  if key == nil then
    error('lkp: key passed to file is nil in ' .. tostring(hpath))
  end
  if hpath == nil or hpath == '' then
    error('lkp: file path is invalid or empty')
  end
  if inner == nil then
    inner = false
  end

  if hashTables[hpath] == nil then
    local success, dataEntries = safeLoadFileSandbox(hpath)
    if not success then
      error('lkp: failed to safely load file at ' .. tostring(hpath) .. ': ' .. tostring(dataEntries))
    end
    if type(dataEntries) ~= 'table' then
      error('lkp: invalid format in file at ' .. tostring(hpath))
    end

    local tbl = {}
    for _, entry in ipairs(dataEntries) do
      local keys = entry[1]
      local value = entry[2]

      if value == nil then
        goto continue
      end

      if type(keys) == 'table' then
        for _, k in ipairs(keys) do
          if type(k) == 'number' then
            k = tostring(k)
          end
          tbl[k] = value
        end
      elseif type(keys) == 'string' or type(keys) == 'number' then
        if type(keys) == 'number' then
          keys = tostring(keys)
        end
        tbl[keys] = value
      else
        goto continue
      end
      ::continue::
    end
    hashTables[hpath] = tbl
  end

  local val = hashTables[hpath][key]
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
------------------------------------------------------------------
function lvars(filePath)
  local res = {result = '', skip = true}

  if filePath == nil or filePath == '' then
    error('lvars: file path is invalid or empty')
    return res
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

  if resultTable == nil then
    resultTable = res
  else
    if resultTable.result == nil then
      resultTable.result = res.result
    end
    if resultTable.skip == nil then
      resultTable.skip = res.skip
    end
  end
  return resultTable
end

)";
static const size_t luaSourceSize = sizeof(luaSourceCode);

#endif // LUA_EMBEDDED_H
