// Validator JSON Schema ringkas untuk memeriksa input alat dari AI sebelum dijalankan.
// Mendukung: type, required, properties, items, enum, additionalProperties:false.

function typeOf(v) {
  if (Array.isArray(v)) return 'array';
  if (v === null) return 'null';
  if (Number.isInteger(v)) return 'integer';
  return typeof v;
}

function matchesType(v, t) {
  const actual = typeOf(v);
  if (Array.isArray(t)) return t.some((x) => matchesType(v, x));
  if (t === 'number') return actual === 'number' || actual === 'integer';
  return actual === t;
}

export function validate(schema, value, path = 'input') {
  const errors = [];
  if (!schema) return errors;
  if (schema.type && !matchesType(value, schema.type)) {
    errors.push(`${path}: harus bertipe ${schema.type}`);
    return errors;
  }
  if (schema.enum && !schema.enum.includes(value)) errors.push(`${path}: harus salah satu dari ${schema.enum.join(', ')}`);
  if (typeOf(value) === 'object') {
    for (const key of schema.required || []) {
      if (value[key] === undefined) errors.push(`${path}.${key}: wajib diisi`);
    }
    const props = schema.properties || {};
    for (const [k, v] of Object.entries(value)) {
      if (props[k]) errors.push(...validate(props[k], v, `${path}.${k}`));
      else if (schema.additionalProperties === false) errors.push(`${path}.${k}: properti tidak dikenal`);
    }
  }
  if (typeOf(value) === 'array' && schema.items) {
    value.forEach((v, i) => errors.push(...validate(schema.items, v, `${path}[${i}]`)));
  }
  return errors;
}
