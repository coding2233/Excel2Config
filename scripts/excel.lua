local excel = read_excel("example.xlsx")

-- log.debug(excel)

-- sheet
for kSheet,vSheet in pairs(excel) do
    log.debug(kSheet,vSheet)
    -- row
    for kRow,vRow in pairs(vSheet) do
        log.debug(kRow,vRow)
        -- cell
        for kCell,vCell in pairs(vRow) do
            log.debug(kCell,vCell)
        end
    end
 
end




-- function write_excel(excel_name)
--     local wb,wx = load_excel();

--     -- for 循环
--     write_worksheet(wx,1,2,"xxx")

--     save_excel(wb,"test.xlsx");
-- end

