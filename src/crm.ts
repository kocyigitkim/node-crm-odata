import { HttpsAgent } from "agentkeepalive";
import { URL } from "url";
import { crm_ntlm_auth } from './crm.ad';


//#region Metadata
export class EntityMetadata {
    public Name: string;
    public PrimaryKey: string;
    public Fields: FieldMetadata[];
}
export class FieldMetadata {
    public Name: string;
    public SchemaName: string;
    public Type: string;
    public LookupEntityName: string;
    public LookupEntityPrimaryKey: string;
}
//#endregion

//#region Structures
export class CRMReference {
    public Id: string;
    public LogicalName: string;

    constructor(initial: CRMReference|any) {
        this.Id = initial.Id;
        this.LogicalName = initial.LogicalName;
    }
}
export class CRMField {
    public Name: string;
    public Value: Object;
    public Entity: CRMEntity;
    public Type: string;

    constructor(initial: CRMField | any) {
        this.Name = initial.Name;
        this.Value = initial.Value;
        this.Entity = initial.Entity;
        this.Type = FindFieldType(initial.Value);
    }
    has() {
        return this.Value != undefined && this.Value != null;
    }
    set(value: Object) {
        this.Value = ConvertFieldType(value, this.Type);
    }
    get(): any {
        return this.Value;
    }
}
export class CRMEntity {
    public Fields: CRMField[] = [];
    public LogicalName: string = null;
    public EntityId: string = null;
    get EntityReference(): CRMReference {
        return new CRMReference({ Id: this.EntityId, LogicalName: this.LogicalName });
    }
    constructor(LogicalName?: string, initial?: Object) {
        this.LogicalName = LogicalName;
        if (initial) {
            for (var p in initial) {
                this.set(p, initial[p]);
            }
        }
    }
    has(name: string): Boolean {
        return Boolean(this.Fields.find(f => f.Name == name));
    }
    get(name: string): CRMField {
        return this.Fields.find(f => f.Name == name);
    }
    set(name: string, value: Object): any {
        if (this.has(name)) {
            this.Fields.find(f => f.Name == name).Value = value;
        }
        else {
            this.Fields.push(new CRMField({ Name: name, Value: value, Entity: this }));
        }
    }
    remove(name: string) {
        this.Fields = this.Fields.filter(f => f.Name != name);
    }
    fill(data: Object) {
        for (var p in data) {
            if (p.indexOf("@") > -1) continue;
            var v = data[p];
            if (p.startsWith("_") && p.endsWith("_value")) {
                var pname = p.substring(1, p.length - 6);
                var entityName = data[p + "@Microsoft.Dynamics.CRM.lookuplogicalname"];
                this.set(pname, new CRMReference({ Id: v, LogicalName: entityName }));
            }
            else {
                if (data[p + "_base"] && data[p + "@OData.Community.Display.V1.FormattedValue"] && data[p + "_base@OData.Community.Display.V1.FormattedValue"]) {
                    this.set(p, typeof v === "string" ? parseFloat(v) : v);
                }
                else {
                    this.set(p, v);
                }
            }
        }
    }
    toJson(manager: CRMManager): any {
        var entityDef = manager.metadata.filter(p => p.Name == this.LogicalName)[0];
        var result: any = {};
        for (var f of this.Fields) {
            var fieldType = f.Type;
            var fieldVal = f.get();
            if (fieldType === "Lookup" && entityDef) {
                var lookupEntityName = fieldVal.LogicalName;
                var lookupEntityPrimaryKey = lookupEntityName + "id";
                var targetFields = entityDef.Fields.filter(p => p.LookupEntityName == lookupEntityName && p.LookupEntityPrimaryKey == lookupEntityPrimaryKey);
                var targetField = targetFields[0];
                if (targetFields.length > 1) {
                    var fieldNameWithoutId = f.Name.endsWith("id") ? f.Name.substring(0, f.Name.length - 2) : f.Name;
                    targetField = targetFields.filter(p => p.Name == fieldNameWithoutId)[0];
                }
                var fieldName = f.Name;
                if (targetField) {
                    fieldName = targetField.SchemaName;
                }
                if (fieldVal.Id) {
                    result[fieldName + "@odata.bind"] = `/${getPluralName(lookupEntityName)}(${formatGuid(fieldVal.Id)})`;
                    continue;
                }
            }
            result[f.Name] = ConvertFieldType(fieldVal, fieldType);
        }
        return result;
    }
    cloneBy(fieldNames: string[]): CRMEntity {
        var result = new CRMEntity(this.LogicalName);
        result.EntityId = this.EntityId;
        for (var f of this.Fields) {
            if (fieldNames.indexOf(f.Name) > -1) {
                result.set(f.Name, f.get());
            }
        }
        return result;
    }
}
//#endregion

export class CRMConnection {
    public url: string;
    public username: string;
    public password: string;
    public domain: string;

    constructor(initial: CRMConnection | any) {
        for (var k in initial) this[k] = initial[k];
        if (this.url.endsWith("/")) this.url = this.url.substr(0, this.url.length - 1);
    }
}

function getPluralName(name: string): string {
    if (name.endsWith("y")) {
        return name.substr(0, name.length - 1) + "ies";
    }
    else if (name.endsWith("s") || name.endsWith("x")) {
        return name + "es";
    }
    else {
        return name + "s";
    }
    return name;
}
function FindFieldType(value: any): string {
    if (value instanceof CRMReference) return "Lookup";

    if (Array.isArray(value)) {
        return "Lookup";
    }
    else {
        if (value === true || value === false) return "TwoOptions";
        if (typeof value === "number") return "Decimal";

        const lookupPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
        if (lookupPattern.test(value)) {
            return "Lookup";
        }

        try {
            var dt = new Date(value.toString());
            if (dt.toString() != "Invalid Date") return "DateTime";
        } catch (err) { }

        return "String";
    }
}
function ConvertFieldType(value: any, type: string): any {
    if (value === null || value === undefined) return null;
    if (type === "Lookup") {
        return value;
    }
    else if (type === "TwoOptions") {
        return value === "true" || value === true;
    }
    else if (type === "Decimal") {
        return parseFloat(value);
    }
    else if (type === "DateTime") {
        return new Date(value);
    }
    else {
        return value.toString();
    }
}


function InitializeMetadata(rawBody: any): EntityMetadata[] {
    const xml2js = require('xml2json');
    var body = xml2js.toJson(rawBody, { object: true });
    var entities = [];
    for (var item of body["edmx:Edmx"]["edmx:DataServices"].Schema.EntityType) {
        var entity = new EntityMetadata();
        entity.Name = item.Name;
        var fields = [];
        if (item.NavigationProperty) {
            if (!Array.isArray(item.NavigationProperty)) item.NavigationProperty = [item.NavigationProperty];

            for (var subitem of item.NavigationProperty) {
                var field = new FieldMetadata();
                field.SchemaName = subitem.Name;
                field.Type = subitem.Type;
                if (field.SchemaName.startsWith("_") && field.SchemaName.endsWith("_value")) {
                    field.SchemaName = field.SchemaName.substr(1, field.SchemaName.length - 9);
                }
                if (subitem.ReferentialConstraint) {
                    var subitemConst = subitem.ReferentialConstraint;
                    field.Name = subitemConst.Property;
                    if (field.Name.startsWith("_") && field.Name.endsWith("_value")) {
                        field.Name = field.Name.substr(1, field.Name.length - 9);
                    }
                    field.LookupEntityName = subitem.Type.replace("mscrm.", "");
                    field.LookupEntityPrimaryKey = subitemConst.ReferencedProperty;
                }
                else {
                    field.Name = field.SchemaName;
                }
                fields.push(field);
            }
        }
        entity.Fields = fields;
        entities.push(entity);
    }

    return entities;
}
function formatGuid(guid: string): string {
    return guid.replace(/[{}]/g, "").toLowerCase();
}
export interface RequestCallBackDelegate {
    (req: any): any;
}
export interface ResponseCallbackDelegate {
    (res: any): any;
}
/**
    @member {EntityMetadata[]} metadata
*/
export class CRMManager {
    public Connection: CRMConnection;
    public OnError: Function[];
    public OnConnectionError: Function[];
    public agent: HttpsAgent;
    public headers: any;
    public metadata: EntityMetadata[];

    constructor(connection: CRMConnection) {
        this.Connection = connection;
        this.OnError = [];
        this.OnConnectionError = [];
    }
    async SendRequestAsync(path: string, requestCallback: RequestCallBackDelegate, responseCallback: ResponseCallbackDelegate): Promise<any> {
        var callError = null;
        var middleRequestCallback = ((requestCallback, req) => {
            return requestCallback(req);
        }).bind(this, requestCallback);
        var result: any = await crm_ntlm_auth(this.agent, this.Connection.url + path, this.Connection.username, this.Connection.password, this.Connection.domain, new URL(this.Connection.url).hostname, middleRequestCallback).catch((err) => callError = err);
        if (!callError && result.statusCode >= 200 && result.statusCode < 300) {
            if (result.headers["odata-entityid"] && result.body.length === 0) {
                return responseCallback({ "id": result.headers["odata-entityid"].split("(")[1].split(")")[0] });
            }
            else if (result.body.length === 0) return responseCallback(true);
            return responseCallback(JSON.parse(result.body));
        }
        else {
            this.OnError.forEach((callback) => callback(callError || new Error(result.body)));
        }
    }

    async ConnectAsync(): Promise<Boolean> {
        var callError = null;
        var HttpsAgent = require('agentkeepalive').HttpsAgent;
        this.agent = new HttpsAgent();
        var result: any = await crm_ntlm_auth(this.agent, this.Connection.url + "/$metadata#EntityDefinitions/Attributes", this.Connection.username, this.Connection.password, this.Connection.domain, new URL(this.Connection.url).hostname).catch((err) => callError = err);
        if (!callError && result.statusCode >= 200 && result.statusCode < 300) {
            this.headers = result.headers;
            const skipHeaders = ["content-type", "content-length", "connection"];
            for (var k of skipHeaders) {
                if (this.headers[k]) {
                    delete this.headers[k];
                }
            }
            this.metadata = InitializeMetadata(result.body);
            return true;
        }
        else {
            this.OnError.forEach((callback) => callback(callError || new Error(result.body)));
        }
        return false;
    }

    async RetrieveAsync(ref: CRMReference): Promise<CRMEntity> {
        var results = await this.SendRequestAsync(`/${getPluralName(ref.LogicalName)}(${formatGuid(ref.Id)})`, (req) => {
            req.method = "get";
            return req;
        }, async (body) => {
            if (body) {
                var entity = new CRMEntity();
                entity.LogicalName = ref.LogicalName;
                entity.EntityId = ref.Id;
                entity.fill(body);

                return entity;
            }
        });
        return results;
    }
    async RetrieveQueryAsync(logicalName: string, query: string): Promise<CRMEntity[]> {
        var results = await this.SendRequestAsync(`/${getPluralName(logicalName)}?${query}`, (req) => {
            req.method = "get";
            req.headers.Prefer = 'odata.include-annotations="*"';
            return req;
        }, async (body) => {
            if (body && body.value) {
                var entities = [];
                for (var v of body.value) {
                    var entity = new CRMEntity();
                    entity.LogicalName = logicalName;
                    entity.fill(v);
                    var entityId = entity.get(logicalName + "id");
                    if (entityId) {
                        entity.EntityId = entityId.get();
                        entity.remove(logicalName + "id");
                    }
                    entities.push(entity);
                }
                return entities;
            }
        });
        return results;
    }
    async RetrieveFetchXmlAsync(fetchXml: string): Promise<CRMEntity[]> {
        var entityName = fetchXml.match(/<(e|E)ntity (n|N)ame="([^"]+)"/)[3];

        var results = await this.SendRequestAsync(`/${getPluralName(entityName)}`, (req) => {
            req.method = "get";
            req.headers.Prefer = 'odata.include-annotations="*"';
            req.headers.FetchXml = fetchXml.replace(/[\n\r]/g, "");
            return req;
        }, async (body) => {
            if (body && body.value) {
                var entities = [];
                for (var v of body.value) {
                    var entity = new CRMEntity();
                    entity.LogicalName = entityName;
                    entity.fill(v);
                    var entityId = entity.get(entityName + "id");
                    if (entityId) {
                        entity.EntityId = entityId.get();
                        entity.remove(entityName + "id");
                    }
                    entities.push(entity);
                }
                return entities;
            }
        });
        return results;
    }
    async CreateAsync(entity: CRMEntity): Promise<Boolean> {
        var results = await this.SendRequestAsync(`/${getPluralName(entity.LogicalName)}`, (req) => {
            req.method = "post";
            req.headers["Content-Type"] = "application/json";
            req.body = JSON.stringify(entity.toJson(this));
            return req;
        }, async (body) => {
            if (body) {
                entity.EntityId = body.id;
                return true;
            }
            return false;
        });
        return results;
    }

    async UpdateAsync(entity: CRMEntity): Promise<Boolean> {
        var results = await this.SendRequestAsync(`/${getPluralName(entity.LogicalName)}(${formatGuid(entity.EntityId)})`, (req) => {
            req.method = "patch";
            req.headers["Content-Type"] = "application/json";
            req.body = JSON.stringify(entity.toJson(this));
            return req;
        }, async (body) => {
            if (body && body.id) {
                return true;
            }
            return false;
        });
        return results;
    }
    async SetStateAsync(ref: CRMReference, state: Number, status: Number) : Promise<Boolean> {
        var results = await this.SendRequestAsync(`/${getPluralName(ref.LogicalName)}(${formatGuid(ref.Id)})`, (req) => {
            req.method = "patch";
            req.headers["Content-Type"] = "application/json";
            req.body = JSON.stringify({
                statecode: state,
                statuscode: status
            });
            return req;
        }, async (body) => {
            if (body && body.id) {
                return true;
            }
            return false;
        });
        return results;
    }
    async DeleteAsync(ref: CRMReference): Promise<Boolean> {
        var results = await this.SendRequestAsync(`/${getPluralName(ref.LogicalName)}(${formatGuid(ref.Id)})`, (req) => {
            req.method = "delete";
            return req;
        }, async (body) => {
            if (body === true) {
                return true;
            }
            return false;
        });
        return results;
    }
    async AssociateAsync(source: CRMReference, dest: CRMReference, relationShipName: string): Promise<Boolean> {
        var results = await this.SendRequestAsync(`/${getPluralName(source.LogicalName)}(${formatGuid(source.Id)})/${relationShipName}/$ref`, (req) => {
            req.method = "post";
            req.headers["Content-Type"] = "application/json";
            req.body = JSON.stringify({
                "@odata.id": `${this.Connection.url}/${getPluralName(dest.LogicalName)}(${formatGuid(dest.Id)})`
            });
            return req;
        }, async (body) => {
            if (body === true) {
                return true;
            }
            return false;
        });
        return results;
    }
    async CreateOrUpdateAsync(entity: CRMEntity) : Promise<Boolean> {
        if (entity.EntityId && entity.EntityId.replace(/[0\-]/g, '').length > 0) {
            return await this.UpdateAsync(entity);
        }
        else {
            return await this.CreateAsync(entity);
        }
    }
}
