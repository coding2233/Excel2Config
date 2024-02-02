.textpro 注意事项(与json的区别)
```json
# textproto
CardList: [
    {
        id: 100001,
        name: "Sunbreaker",
        name_en: "Sunbreaker",
        name_key: "faction_1_unit_sunbreaker_name",
        hide_in_collection: false,
        faction: 1,
        race_ids: -1,
        races: "unit",
        rarity: 2,
        attack: 2,
        max_hp: 4,
        mana_cost: 4,
        img: "units/f1_sunbreaker.png",
        fx: "FX.Cards.Faction1.Sunriser",
        copies: 3,
        desc: "力场。你将军的血源法术变为 Tempest。",
        sounds: {
            key: "apply"
            value: "sfx_f1_sunriser_death_alt.m4a"
        },
        sounds: {
            key: "walak"
            value: "sfx_f1_sunriser_death_alt.m4a"
        },
        animations: [
            "breathing",
            "idle",
            "walk",
            "attack",
            "damage",
            "death"
        ],
        attack_delay: 0.6,
        attack_release_delay: 0
    }
]
```

1. 没有最外层的大括号`{}`
2. key值没有双引号`""`
3. map格式示例  
```json
# json
sound: {
    "apply": "sfx_ui_booster_packexplode.m4a",
    "walk": "sfx_neutral_ladylocke_attack_impact.m4a",
    "attack": "sfx_f1_sunriser_attack_swing.m4a",
    "receiveDamage": "sfx_f1_sunriser_hit_noimpact.m4a",
    "attackDamage": "sfx_f1_sunriser_attack_impact.m4a",
    "death": "sfx_f1_sunriser_death_alt.m4a"
}

#textproto
    sounds: {
        key: "apply"
        value: "sfx_ui_booster_packexplode.m4a"
    },
    sounds: {
        key: "walk"
        value: "sfx_neutral_ladylocke_attack_impact.m4a"
    },
    sounds: {
        key: "attack"
        value: "sfx_f1_sunriser_attack_swing.m4a"
    },
    sounds: {
        key: "receiveDamage"
        value: "sfx_f1_sunriser_hit_noimpact.m4a"
    },
    sounds: {
        key: "attackDamage"
        value: "sfx_f1_sunriser_attack_impact.m4a"
    },
    sounds: {
        key: "death"
        value: "sfx_f1_sunriser_death_alt.m4a"
    },
```

textpro转字节码 `--encode`

```bash
protoc.exe --proto_path=. --encode=Wanderer.Config.Collect Collect.xlsx.proto < collect.textproto > collect.bytes
```

textpro转文本 `--decode`
```bash
protoc.exe --proto_path=. --decode=Wanderer.Config.Collect Collect.xlsx.proto < collect.bytes
```


```bash
#打包成单文件
dotnet publish -f net6.0 -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
```