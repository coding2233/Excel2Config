print("hello xlnt lua.")

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