local excel_pb = {}

local pb = require "pb"
local protoc = require "protoc"

local collect={
    CardList={
        {id=100001,name="Sunbreaker",name_en="Sunbreaker",name_key="faction_1_unit_sunbreaker_name",hide_in_collection=false,attack=2,max_hp=4,mana_cost=4,img="units/f1/f1_sunbreaker.png",fx="FX.Cards.Faction1.Sunriser",copies=3,desc="力场。你将军的血源法术变为 Tempest。",animations={"breathing","idle","walk","attack","damage","death",},attack_delay=0.6,attack_release_delay=0,},
        {id=100002,name="Scintilla",name_en="Scintilla",name_key="faction_1_unit_scintilla_name",hide_in_collection=false,attack=3,max_hp=4,mana_cost=3,img="units/f1/f1_scintilla.png",fx="FX.Cards.Faction1.Scintilla",copies=3,desc="血涌：回复你的将军 3 点生命值。",animations={"breathing","idle","walk","attack","damage","death",},attack_delay=0.6,attack_release_delay=0,},
        {id=100003,name="Excelsious",name_en="Excelsious",name_key="faction_1_unit_excelsious_name",hide_in_collection=false,attack=6,max_hp=6,mana_cost=8,img="units/f1/f1_excelsious.png",fx="FX.Cards.Faction1.IroncliffeGuardian",copies=3,desc="挑衅。迅捷。本场赛局中任何单位受到治疗都会使其 +1/+1。",animations={"breathing","idle","walk","attack","damage","death",},attack_delay=0.7,attack_release_delay=0,},
    }
}

local proto_scm = [[
    syntax = "proto3";

    message Card
    {
            //唯一标识
            int32 id = 1;
            //名称
            string name = 2;
            //英文原名
            string name_en = 3;
            //名称键值
            string name_key = 4;
    }

    message Collect
    {
            //列表
            repeated Card CardList = 1;
    }
]]

function excel_pb.ProtobufExcelEncode(proto,proto_name,data_table)
    -- data_table = collect
    -- proto = proto_scm
    log.debug(string.format("ProtobufExcelEncode\n %s\n%s\n%s",proto_name,proto,data_table))
    -- log.debug(proto)
    assert(protoc:load(proto),"protoc:load error")
    local bytes = assert(pb.encode(proto_name,data_table))
    -- log.debug("protobuf_encode pb.encode",bytes)
    log.debug(pb.tohex(bytes))

    return bytes
end





function PBTest()
    -- load schema from text (just for demo, use protoc.new() in real world)
    assert(protoc:load [[
        message Phone {
        optional string name        = 1;
        optional int64  phonenumber = 2;
        }
        message Person {
        optional string name     = 1;
        optional int32  age      = 2;
        optional string address  = 3;
        repeated Phone  contacts = 4;
        } ]])
    
    -- lua table data
    local data = {
        name = "ilse",
        age  = 18,
        contacts = {
        { name = "alice", phonenumber = 12312341234 },
        { name = "bob",   phonenumber = 45645674567 }
        }
    }
    
    -- encode lua table data into binary format in lua string and return
    local bytes = assert(pb.encode("Person", data))
    log.debug(pb.tohex(bytes))
    
    -- and decode the binary data back into lua table
    local data2 = assert(pb.decode("Person", bytes))
    log.debug(require "serpent".block(data2))
end


-- function protobuf_decode(proto,proto_name,data_bytes)

--     assert(protoc:load(proto))

--     local data = assert(pb.decode(proto_name,data_bytes))
--     -- log.debug(pb.tohex(bytes))
-- end

return excel_pb