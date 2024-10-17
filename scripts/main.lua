require("config")

print("hello xlnt lua.")

local configs = parse_args()

for k,v in pairs(configs) do
    print(k,v)
end

-- function test_require()
--     local pb = require("pb")
--     print(pb)
-- end

-- local status,error = pcall(test_require)
-- print(error)

local excel = read_excel("example.xlsx")

-- print(excel)

-- sheet
for kSheet,vSheet in pairs(excel) do
    print(kSheet,vSheet)
    -- row
    for kRow,vRow in pairs(vSheet) do
        print(kRow,vRow)
        -- cell
        for kCell,vCell in pairs(vRow) do
            print(kCell,vCell)
        end
    end
 
end

print("[end] xlnt lua")
-- local test_reult = test_register();
-- print(test_reult)



-- function write_excel(excel_name)
--     local wb,wx = load_excel();

--     -- for 循环
--     write_worksheet(wx,1,2,"xxx")

--     save_excel(wb,"test.xlsx");
-- end

