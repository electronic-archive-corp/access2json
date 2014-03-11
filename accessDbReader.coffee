###
    Класс для прохода по нужным данным

    Потомок должен определить метод
    process: (state, level, data)
        state - текущее состояние в потоке данных
        level - текущий уровень вложенности данных
        data - данные доступные на этом уровне (при необходимости бОльших или других данных - менять нужно именно в этом классе)

    Как ЭТО работает:
        ключевой метод - точка старта - read (dbPath, Config)
        Подключение к базе
        Считывается первый query из конфига
        Выбираются данные из базы
        По мере выборки данных дергается метод process
        Проходим только по полям которые определены в конфиге - из них и планируется составить будующий result
###
class @AccessDbReader
    ADBR_CREATED: 1
    ADBR_STARTED: 2
    ADBR_ARRAY_START: 3
    ADBR_OBJ_START: 4
    ADBR_NEXT_OBJ: 5
    ADBR_ERROR: 6
    ADBR_OBJ_END: 7
    ADBR_ARRAY_END: 8
    ADBR_END: 9
    ADBR_PROP_NAME: 10
    ADBR_PROP_VALUE: 11
    ADBR_NEXT_PROP: 12
    ADBR_BEFORE_CHILD: 13
    ADBR_AFTER_CHILD: 14

    @conn: null
    @curLevel = 0

    constructor: ()->
        @process(@ADBR_CREATED, 0)


    process: (state, level, data)->
        #override


    template2string: (t, data) ->
        if typeof (t) != "string"
            return t
        start = t.indexOf("{{")
        if start < 0
            return t
        start = start + 2
        end = t.indexOf("}}", start)
        if end < 0
            return t
        end = end + 2

        #ok, we have variable
        left = t.substr(0, start - 2)
        right = t.substr(end)
        t = data[t.substr(start, end - start - 2)]
        if typeof (t) is "string"
            t = t.replace(/\"/g, "&#34;")
            t = t.replace(/\'/g, "&#39;")
        if left.length is 0 and right.length is 0
            return t

        t = left + t + right
        return @template2string(t, data)


    construct: (config, data) ->
        SQL = @template2string(config.query, data)
        rs = new ActiveXObject("ADODB.Recordset")
        adOpenDynamic = 2
        adLockOptimistic = 3
        rs.open SQL, @conn, adOpenDynamic, adLockOptimistic
        @process(@ADBR_ARRAY_START, @curLevel, data)
        @curLevel++;
        if rs.Fields.Count
            if not rs.bof and not rs.eof
                rs.MoveFirst()
                first = true
                until rs.eof
                    if !first
                        @process(@ADBR_NEXT_OBJ, @curLevel, data)
                    dbRow = {}
                    x = 0
                    while x < rs.Fields.Count
                        dbRow[rs.Fields(x).Name] = rs.Fields(x).Value
                        x++

                    @process(@ADBR_OBJ_START, @curLevel, data)
                    firstProp = true
                    for i of config.template
                        unless config.template.hasOwnProperty(i)
                            continue
                        unless firstProp
                            @process(@ADBR_NEXT_PROP, @curLevel, data)
                        @process(@ADBR_PROP_NAME, @curLevel, i)

                        field = config.template[i]
                        if typeof (field) is "object"
                            @curLevel++;
                            @process @ADBR_BEFORE_CHILD, @curLevel, data
                            @construct field, dbRow
                            @process @ADBR_AFTER_CHILD, @curLevel, data
                            @curLevel--;
                        else
                            @process @ADBR_PROP_VALUE, @curLevel, @template2string(field, dbRow)
                        firstProp = false
                    @process(@ADBR_OBJ_END, @curLevel, data)
                    first = false
                    rs.MoveNext()
        @curLevel--;
        @process(@ADBR_ARRAY_END, @curLevel)
        rs.close()


    read: (dbPath, Config) ->
        cs = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " + dbPath + ";Jet OLEDB:Engine Type=4;Persist Security Info = false"
        @conn = new ActiveXObject("ADODB.Connection")
        @conn.open cs, "", "" # connection string, user, pass
        @curLevel = 0
        @process(@ADBR_STARTED, 0, Config);
        @construct Config, {}
        @process(@ADBR_END, 0, Config);
        @conn.close()

