local pb = require "pb"
local protoc = require "protoc"

function protobuf_encode(proto,proto_name,data_table)
    assert(protoc:load(proto))
    local bytes = assert(pb.encode(proto_name,data_table))
    print(pb.tohex(bytes))
end


function protobuf_decode(proto,proto_name,data_bytes)

    assert(protoc:load(proto))

    local data = assert(pb.decode(proto_name,data_bytes))
    -- print(pb.tohex(bytes))
end