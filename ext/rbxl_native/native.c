/*
 * rbxl_native - Optional C extension for rbxl
 *
 * Parses xlsx sheet XML via libxml2 SAX2 directly, bypassing
 * Nokogiri's per-node Ruby object allocation overhead.
 *
 * Security considerations:
 *   - All buffers are dynamically allocated and grown as needed (no fixed limits)
 *   - Shared string index is bounds-checked
 *   - XML parser depth is limited to prevent XML bomb attacks
 *   - Parser context is cleaned up even when Ruby exceptions occur
 *   - All string inputs are validated
 */

#include <ruby.h>
#include <ruby/encoding.h>
#include <libxml/parser.h>
#include <libxml/SAX2.h>
#include <string.h>
#include <stdlib.h>

static rb_encoding *enc_utf8;

static inline VALUE make_utf8_str(const char *ptr, long len)
{
    VALUE s = rb_str_new(ptr, len);
    rb_enc_associate(s, enc_utf8);
    return s;
}

#define INITIAL_BUF_CAP 256
#define MAX_XML_DEPTH   64
#define MAX_TOTAL_BYTES (512 * 1024 * 1024) /* 512 MB hard limit on accumulated text */

/* ------------------------------------------------------------------ */
/* Dynamic buffer                                                      */
/* ------------------------------------------------------------------ */

typedef struct {
    char  *data;
    size_t len;
    size_t cap;
} dynbuf;

static void dynbuf_init(dynbuf *b)
{
    b->data = NULL;
    b->len  = 0;
    b->cap  = 0;
}

static void dynbuf_free(dynbuf *b)
{
    if (b->data) { xfree(b->data); b->data = NULL; }
    b->len = b->cap = 0;
}

static void dynbuf_clear(dynbuf *b)
{
    b->len = 0;
}

static int dynbuf_append(dynbuf *b, const char *src, size_t n)
{
    if (n == 0) return 1;
    size_t needed = b->len + n;
    if (needed > MAX_TOTAL_BYTES) return 0; /* refuse oversized input */
    if (needed > b->cap) {
        size_t newcap = b->cap ? b->cap : INITIAL_BUF_CAP;
        while (newcap < needed) newcap *= 2;
        char *tmp = xrealloc(b->data, newcap);
        b->data = tmp;
        b->cap  = newcap;
    }
    memcpy(b->data + b->len, src, n);
    b->len += n;
    return 1;
}

/* ------------------------------------------------------------------ */
/* Parse context                                                       */
/* ------------------------------------------------------------------ */

typedef struct {
    /* Row / cell counters */
    int row_count;
    int cell_count;

    /* Nesting state */
    int in_row;
    int in_cell;
    int collecting;  /* currently inside <v> or <t> */
    int is_v;        /* distinguishes </v> from </t> */
    int depth;       /* current nesting depth */

    /* Cell type attribute value ('s', 'b', 'n', ...) */
    char cell_type;
    int  has_cell_type;

    /* Buffers */
    dynbuf text_buf; /* accumulates character data for current <v>/<t> */
    dynbuf raw_buf;  /* accumulated raw value for current cell */
    int    has_raw;

    /* Cell coordinate (for full read mode) */
    dynbuf cell_ref; /* "r" attribute of <c> */

    /* Row index (for full read mode) */
    int row_index;

    /* Shared strings (Ruby Array, must be marked during GC) */
    VALUE shared_strings;
    long  shared_strings_len;

    /* Current row (Ruby Array) */
    VALUE current_row;

    /* Mode: 0 = values_only, 1 = full (ReadOnlyCell + Row) */
    int full_mode;

    /* Ruby classes for full mode (looked up once at init) */
    VALUE cReadOnlyCell;
    VALUE cRow;

    /* Error flag — set if a callback wants to abort */
    int error;
    char error_msg[256];
} parse_ctx;

/* ------------------------------------------------------------------ */
/* Value coercion (mirrors Rbxl::ReadOnlyWorksheet#coerce_value)       */
/* ------------------------------------------------------------------ */

static VALUE coerce_value(parse_ctx *c)
{
    if (!c->has_raw) {
        if (c->has_cell_type && c->cell_type == 'b') return Qfalse;
        return Qnil;
    }
    /* inlineStr with empty <t></t> should return "" not nil */
    if (c->raw_buf.len == 0) {
        if (c->has_cell_type && c->cell_type == 'i') return make_utf8_str("", 0);
        if (c->has_cell_type && c->cell_type == 'b') return Qfalse;
        return Qnil;
    }

    const char *raw = c->raw_buf.data;
    size_t len = c->raw_buf.len;

    if (c->has_cell_type) {
        switch (c->cell_type) {
        case 's': { /* shared string index */
            long idx = 0;
            for (size_t i = 0; i < len; i++) {
                unsigned char ch = (unsigned char)raw[i];
                if (ch < '0' || ch > '9') {
                    /* malformed index — return raw string */
                    return make_utf8_str(raw, (long)len);
                }
                long next = idx * 10 + (ch - '0');
                if (next < idx) { /* overflow */
                    return make_utf8_str(raw, (long)len);
                }
                idx = next;
            }
            if (idx < 0 || idx >= c->shared_strings_len) {
                /* out of bounds — return raw string rather than crashing */
                return make_utf8_str(raw, (long)len);
            }
            return rb_ary_entry(c->shared_strings, idx);
        }
        case 'b': /* boolean */
            return (len == 1 && raw[0] == '1') ? Qtrue : Qfalse;
        case 'i': /* inlineStr — raw is the text content */
            return make_utf8_str(raw, (long)len);
        default:
            /* "str" and other text types — return as-is */
            if (c->cell_type != 'n')
                return make_utf8_str(raw, (long)len);
            break;
        }
    }

    /* Infer numeric scalar */
    int has_dot = 0;
    size_t start = 0;
    if (raw[0] == '-') { start = 1; if (len == 1) return make_utf8_str(raw, (long)len); }

    int has_digit = 0;
    for (size_t i = start; i < len; i++) {
        unsigned char ch = (unsigned char)raw[i];
        if (ch >= '0' && ch <= '9') {
            has_digit = 1;
        } else if (ch == '.') {
            if (has_dot) return make_utf8_str(raw, (long)len);
            has_dot = 1;
        } else {
            return make_utf8_str(raw, (long)len);
        }
    }
    if (!has_digit) return make_utf8_str(raw, (long)len);

    /* NUL-terminate for strtod/strtol (buffer always has room) */
    dynbuf_append(&c->raw_buf, "\0", 1);
    c->raw_buf.len--; /* don't count NUL in logical length */

    if (has_dot) {
        return DBL2NUM(strtod(c->raw_buf.data, NULL));
    } else {
        return LONG2NUM(strtol(c->raw_buf.data, NULL, 10));
    }
}

/* ------------------------------------------------------------------ */
/* SAX2 callbacks                                                      */
/* ------------------------------------------------------------------ */

static void on_start_element(void *ctx, const xmlChar *localname,
                             const xmlChar *prefix, const xmlChar *URI,
                             int nb_namespaces, const xmlChar **namespaces,
                             int nb_attributes, int nb_defaulted,
                             const xmlChar **attributes)
{
    parse_ctx *c = (parse_ctx *)ctx;
    (void)prefix; (void)URI; (void)nb_namespaces; (void)namespaces;
    (void)nb_defaulted;

    c->depth++;
    if (c->depth > MAX_XML_DEPTH) {
        c->error = 1;
        snprintf(c->error_msg, sizeof(c->error_msg),
                 "XML depth exceeds limit (%d)", MAX_XML_DEPTH);
        xmlStopParser(NULL); /* will be caught by parse loop */
        return;
    }

    const char *name = (const char *)localname;

    if (name[0] == 'r' && name[1] == 'o' && name[2] == 'w' && name[3] == '\0') {
        c->in_row = 1;
        c->current_row = rb_ary_new();
        if (c->full_mode) {
            c->row_index = 0;
            /* extract "r" attribute for row index */
            for (int i = 0; i < nb_attributes; i++) {
                const char *aname = (const char *)attributes[i * 5];
                if (aname[0] == 'r' && aname[1] == '\0') {
                    const char *vstart = (const char *)attributes[i * 5 + 3];
                    const char *vend   = (const char *)attributes[i * 5 + 4];
                    for (const char *p = vstart; p < vend; p++) {
                        c->row_index = c->row_index * 10 + (*p - '0');
                    }
                    break;
                }
            }
        }
    } else if (name[0] == 'c' && name[1] == '\0') {
        c->in_cell = 1;
        c->has_cell_type = 0;
        c->has_raw = 0;
        dynbuf_clear(&c->raw_buf);
        if (c->full_mode) dynbuf_clear(&c->cell_ref);
        /* extract attributes from the SAX2 attribute array */
        for (int i = 0; i < nb_attributes; i++) {
            const char *aname = (const char *)attributes[i * 5];
            if (aname[0] == 't' && aname[1] == '\0') {
                const char *vstart = (const char *)attributes[i * 5 + 3];
                c->cell_type = vstart[0];
                c->has_cell_type = 1;
            } else if (c->full_mode && aname[0] == 'r' && aname[1] == '\0') {
                const char *vstart = (const char *)attributes[i * 5 + 3];
                const char *vend   = (const char *)attributes[i * 5 + 4];
                dynbuf_append(&c->cell_ref, vstart, (size_t)(vend - vstart));
            }
        }
    } else if (name[0] == 'v' && name[1] == '\0') {
        c->collecting = 1;
        c->is_v = 1;
        dynbuf_clear(&c->text_buf);
    } else if (name[0] == 't' && name[1] == '\0') {
        c->collecting = 1;
        c->is_v = 0;
        dynbuf_clear(&c->text_buf);
    }
}

static void on_end_element(void *ctx, const xmlChar *localname,
                           const xmlChar *prefix, const xmlChar *URI)
{
    parse_ctx *c = (parse_ctx *)ctx;
    (void)prefix; (void)URI;

    if (c->collecting) {
        if (c->is_v) {
            /* end of <v> — copy text_buf to raw_buf */
            dynbuf_clear(&c->raw_buf);
            dynbuf_append(&c->raw_buf, c->text_buf.data, c->text_buf.len);
            c->has_raw = 1;
            c->collecting = 0;
        } else {
            /* end of <t> — append text_buf to raw_buf */
            dynbuf_append(&c->raw_buf, c->text_buf.data, c->text_buf.len);
            c->has_raw = 1;
            c->collecting = 0;
        }
    } else {
        const char *name = (const char *)localname;
        if (name[0] == 'c' && name[1] == '\0') {
            VALUE val = coerce_value(c);
            if (c->full_mode) {
                /* Build ReadOnlyCell.new(coordinate, value) */
                VALUE coord;
                if (c->cell_ref.len > 0) {
                    coord = make_utf8_str(c->cell_ref.data, (long)c->cell_ref.len);
                } else {
                    coord = Qnil;
                }
                VALUE cell = rb_funcall(c->cReadOnlyCell, rb_intern("new"), 2, coord, val);
                rb_ary_push(c->current_row, cell);
            } else {
                rb_ary_push(c->current_row, val);
            }
            c->in_cell = 0;
            c->cell_count++;
        } else if (name[0] == 'r' && name[1] == 'o' && name[2] == 'w' && name[3] == '\0') {
            if (c->full_mode) {
                /* Build Row.new(index: row_index, cells: cells) */
                VALUE kwargs = rb_hash_new();
                rb_hash_aset(kwargs, ID2SYM(rb_intern("index")), INT2NUM(c->row_index));
                rb_hash_aset(kwargs, ID2SYM(rb_intern("cells")), c->current_row);
                VALUE argv[1] = { kwargs };
                VALUE row = rb_funcallv_kw(c->cRow, rb_intern("new"), 1, argv, RB_PASS_KEYWORDS);
                rb_yield(row);
            } else {
                rb_ary_freeze(c->current_row);
                rb_yield(c->current_row);
            }
            c->current_row = Qnil;
            c->in_row = 0;
            c->row_count++;
        }
    }

    c->depth--;
}

static void on_characters(void *ctx, const xmlChar *ch, int len)
{
    parse_ctx *c = (parse_ctx *)ctx;
    if (!c->collecting || len <= 0) return;
    if (!dynbuf_append(&c->text_buf, (const char *)ch, (size_t)len)) {
        c->error = 1;
        snprintf(c->error_msg, sizeof(c->error_msg),
                 "cell text exceeds %d byte limit", MAX_TOTAL_BYTES);
    }
}

/* ------------------------------------------------------------------ */
/* Ensure-style cleanup wrapper                                        */
/* ------------------------------------------------------------------ */

#define IO_READ_CHUNK_BYTES (64 * 1024)

typedef struct {
    parse_ctx     *ctx;
    xmlParserCtxtPtr parser;
    const char    *data;      /* string mode only */
    long           data_len;  /* string mode only */
    VALUE          io;        /* io mode only (Qnil in string mode) */
    long           max_bytes; /* io mode cap; 0 = unbounded */
} parse_args;

static VALUE do_parse(VALUE arg)
{
    parse_args *a = (parse_args *)arg;

    xmlParseChunk(a->parser, a->data, (int)a->data_len, 1 /* terminate */);

    return Qnil;
}

static VALUE do_parse_io(VALUE arg)
{
    parse_args *a = (parse_args *)arg;
    static ID id_read = 0;
    if (!id_read) id_read = rb_intern("read");
    VALUE chunk_size = INT2NUM(IO_READ_CHUNK_BYTES);
    long total = 0;

    while (1) {
        VALUE chunk = rb_funcall(a->io, id_read, 1, chunk_size);
        if (NIL_P(chunk)) break;
        Check_Type(chunk, T_STRING);

        long n = RSTRING_LEN(chunk);
        if (n == 0) break;

        total += n;
        if (a->max_bytes > 0 && total > a->max_bytes) {
            a->ctx->error = 1;
            snprintf(a->ctx->error_msg, sizeof(a->ctx->error_msg),
                     "worksheet bytes exceed limit (%ld)", a->max_bytes);
            break;
        }

        xmlParseChunk(a->parser, RSTRING_PTR(chunk), (int)n, 0);
        if (a->ctx->error) break;
    }

    /* Terminate the parser so any trailing buffered state flushes. */
    xmlParseChunk(a->parser, NULL, 0, 1);
    return Qnil;
}

static VALUE cleanup_parse(VALUE arg)
{
    parse_args *a = (parse_args *)arg;
    if (a->parser) {
        xmlFreeParserCtxt(a->parser);
        a->parser = NULL;
    }
    dynbuf_free(&a->ctx->text_buf);
    dynbuf_free(&a->ctx->raw_buf);
    dynbuf_free(&a->ctx->cell_ref);
    return Qnil;
}

/* ------------------------------------------------------------------ */
/* Common parse setup                                                  */
/* ------------------------------------------------------------------ */

static xmlParserCtxtPtr setup_push_parser(parse_ctx *ctx)
{
    xmlSAXHandler handler;
    memset(&handler, 0, sizeof(handler));
    handler.initialized    = XML_SAX2_MAGIC;
    handler.startElementNs = on_start_element;
    handler.endElementNs   = on_end_element;
    handler.characters     = on_characters;

    xmlParserCtxtPtr parser = xmlCreatePushParserCtxt(
        &handler, ctx, NULL, 0, NULL);

    if (!parser) {
        rb_raise(rb_eRuntimeError, "failed to create libxml2 parser context");
    }

    /* XXE / entity-expansion defense:
     *   - NONET: no network access
     *   - NOENT omitted: user-defined entities are NOT substituted, so
     *     external entities are never resolved and billion-laughs style
     *     expansion cannot trigger. Predefined entities (&amp; etc.) still
     *     reach the characters callback via libxml2's default SAX2 handler.
     *   - HUGE omitted: keep libxml2's built-in parser limits active.
     *   Real xlsx files stay well under these limits (Excel caps cell text
     *   at 32,767 chars), so no throughput loss. */
    xmlCtxtUseOptions(parser, XML_PARSE_NONET);
    return parser;
}

static VALUE run_parse(parse_ctx *ctx, VALUE xml_str)
{
    xmlParserCtxtPtr parser = setup_push_parser(ctx);
    parse_args args = { ctx, parser,
                        RSTRING_PTR(xml_str), RSTRING_LEN(xml_str),
                        Qnil, 0 };

    /* rb_ensure guarantees cleanup even if rb_yield raises */
    rb_ensure(do_parse, (VALUE)&args, cleanup_parse, (VALUE)&args);

    if (ctx->error) {
        rb_raise(rb_eRuntimeError, "rbxl_native: %s", ctx->error_msg);
    }

    return INT2NUM(ctx->row_count);
}

static VALUE run_parse_io(parse_ctx *ctx, VALUE io, long max_bytes)
{
    xmlParserCtxtPtr parser = setup_push_parser(ctx);
    parse_args args = { ctx, parser, NULL, 0, io, max_bytes };

    rb_ensure(do_parse_io, (VALUE)&args, cleanup_parse, (VALUE)&args);

    if (ctx->error) {
        rb_raise(rb_eRuntimeError, "rbxl_native: %s", ctx->error_msg);
    }

    return INT2NUM(ctx->row_count);
}

/* ------------------------------------------------------------------ */
/* Ruby method: Rbxl::Native.parse_sheet(xml_string, shared_strings) */
/* ------------------------------------------------------------------ */

static VALUE rb_native_parse(VALUE self, VALUE xml_str, VALUE shared_strings)
{
    (void)self;
    Check_Type(xml_str, T_STRING);
    Check_Type(shared_strings, T_ARRAY);

    parse_ctx ctx;
    memset(&ctx, 0, sizeof(ctx));
    ctx.shared_strings     = shared_strings;
    ctx.shared_strings_len = RARRAY_LEN(shared_strings);
    ctx.current_row        = Qnil;
    ctx.full_mode          = 0;
    dynbuf_init(&ctx.text_buf);
    dynbuf_init(&ctx.raw_buf);

    return run_parse(&ctx, xml_str);
}

/* ------------------------------------------------------------------ */
/* Ruby method: Rbxl::Native.parse_sheet_full(xml_string, shared_strings) */
/* ------------------------------------------------------------------ */

static VALUE rb_native_parse_full(VALUE self, VALUE xml_str, VALUE shared_strings)
{
    (void)self;
    Check_Type(xml_str, T_STRING);
    Check_Type(shared_strings, T_ARRAY);

    VALUE mRbxl = rb_const_get(rb_cObject, rb_intern("Rbxl"));

    parse_ctx ctx;
    memset(&ctx, 0, sizeof(ctx));
    ctx.shared_strings     = shared_strings;
    ctx.shared_strings_len = RARRAY_LEN(shared_strings);
    ctx.current_row        = Qnil;
    ctx.full_mode          = 1;
    ctx.cReadOnlyCell      = rb_const_get(mRbxl, rb_intern("ReadOnlyCell"));
    ctx.cRow               = rb_const_get(mRbxl, rb_intern("Row"));
    dynbuf_init(&ctx.text_buf);
    dynbuf_init(&ctx.raw_buf);
    dynbuf_init(&ctx.cell_ref);

    return run_parse(&ctx, xml_str);
}

/* ------------------------------------------------------------------ */
/* Ruby method: Rbxl::Native.parse_sheet_io(io, shared_strings, max_bytes) */
/*   Chunk-fed streaming variant of parse_sheet.                        */
/*   max_bytes may be nil to disable the worksheet byte cap.            */
/* ------------------------------------------------------------------ */

static VALUE rb_native_parse_io(VALUE self, VALUE io, VALUE shared_strings, VALUE max_bytes)
{
    (void)self;
    Check_Type(shared_strings, T_ARRAY);

    long max = NIL_P(max_bytes) ? 0 : NUM2LONG(max_bytes);

    parse_ctx ctx;
    memset(&ctx, 0, sizeof(ctx));
    ctx.shared_strings     = shared_strings;
    ctx.shared_strings_len = RARRAY_LEN(shared_strings);
    ctx.current_row        = Qnil;
    ctx.full_mode          = 0;
    dynbuf_init(&ctx.text_buf);
    dynbuf_init(&ctx.raw_buf);

    return run_parse_io(&ctx, io, max);
}

/* ------------------------------------------------------------------ */
/* Ruby method: Rbxl::Native.parse_sheet_full_io(io, shared_strings, max_bytes) */
/* ------------------------------------------------------------------ */

static VALUE rb_native_parse_full_io(VALUE self, VALUE io, VALUE shared_strings, VALUE max_bytes)
{
    (void)self;
    Check_Type(shared_strings, T_ARRAY);

    long max = NIL_P(max_bytes) ? 0 : NUM2LONG(max_bytes);

    VALUE mRbxl = rb_const_get(rb_cObject, rb_intern("Rbxl"));

    parse_ctx ctx;
    memset(&ctx, 0, sizeof(ctx));
    ctx.shared_strings     = shared_strings;
    ctx.shared_strings_len = RARRAY_LEN(shared_strings);
    ctx.current_row        = Qnil;
    ctx.full_mode          = 1;
    ctx.cReadOnlyCell      = rb_const_get(mRbxl, rb_intern("ReadOnlyCell"));
    ctx.cRow               = rb_const_get(mRbxl, rb_intern("Row"));
    dynbuf_init(&ctx.text_buf);
    dynbuf_init(&ctx.raw_buf);
    dynbuf_init(&ctx.cell_ref);

    return run_parse_io(&ctx, io, max);
}

/* ================================================================== */
/* Native writer — generate sheet XML from Ruby Array of Arrays        */
/* ================================================================== */

/* Column name from 1-based index: 1→A, 26→Z, 27→AA, ... */
static void write_column_name(dynbuf *buf, int index)
{
    char tmp[8]; /* max 3 letters for 16384 columns */
    int pos = sizeof(tmp);
    int cur = index;

    while (cur > 0) {
        cur--;
        tmp[--pos] = (char)('A' + (cur % 26));
        cur /= 26;
    }
    dynbuf_append(buf, tmp + pos, sizeof(tmp) - (size_t)pos);
}

static void __attribute__((noinline)) write_int(dynbuf *buf, long val)
{
    char tmp[32];
    int len = snprintf(tmp, sizeof(tmp), "%ld", val);
    dynbuf_append(buf, tmp, (size_t)len);
}

/* XML-escape string and append to buf */
static void write_escaped(dynbuf *buf, const char *str, long slen)
{
    const char *p = str;
    const char *end = str + slen;
    const char *seg_start = p;

    while (p < end) {
        const char *esc = NULL;
        int esc_len = 0;
        switch (*p) {
        case '&':  esc = "&amp;";  esc_len = 5; break;
        case '<':  esc = "&lt;";   esc_len = 4; break;
        case '>':  esc = "&gt;";   esc_len = 4; break;
        case '"':  esc = "&quot;"; esc_len = 6; break;
        }
        if (esc) {
            if (p > seg_start) dynbuf_append(buf, seg_start, (size_t)(p - seg_start));
            dynbuf_append(buf, esc, (size_t)esc_len);
            seg_start = p + 1;
        }
        p++;
    }
    if (seg_start < end) dynbuf_append(buf, seg_start, (size_t)(end - seg_start));
}

/* Write a single cell */
static void write_cell(dynbuf *buf, int col, int row, VALUE value, VALUE cWriteOnlyCell)
{
    #define W(s) dynbuf_append(buf, s, sizeof(s) - 1)

    W("<c r=\"");
    write_column_name(buf, col);
    write_int(buf, row);

    if (rb_obj_is_kind_of(value, cWriteOnlyCell)) {
        VALUE cell_value = rb_funcall(value, rb_intern("value"), 0);
        VALUE style_id   = rb_funcall(value, rb_intern("style_id"), 0);

        W("\"");
        if (!NIL_P(style_id)) {
            W(" s=\"");
            write_int(buf, NUM2LONG(style_id));
            W("\"");
        }

        if (NIL_P(cell_value)) {
            W("/>");
        } else if (cell_value == Qtrue) {
            W(" t=\"b\"><v>1</v></c>");
        } else if (cell_value == Qfalse) {
            W(" t=\"b\"><v>0</v></c>");
        } else if (RB_INTEGER_TYPE_P(cell_value)) {
            W("><v>");
            write_int(buf, NUM2LONG(cell_value));
            W("</v></c>");
        } else if (RB_FLOAT_TYPE_P(cell_value)) {
            W("><v>");
            VALUE fs = rb_funcall(cell_value, rb_intern("to_s"), 0);
            dynbuf_append(buf, RSTRING_PTR(fs), (size_t)RSTRING_LEN(fs));
            W("</v></c>");
        } else {
            VALUE s = rb_funcall(cell_value, rb_intern("to_s"), 0);
            W(" t=\"inlineStr\"><is><t>");
            write_escaped(buf, RSTRING_PTR(s), RSTRING_LEN(s));
            W("</t></is></c>");
        }
        return;
    }

    if (NIL_P(value)) {
        W("\"/>");
    } else if (value == Qtrue) {
        W("\" t=\"b\"><v>1</v></c>");
    } else if (value == Qfalse) {
        W("\" t=\"b\"><v>0</v></c>");
    } else if (RB_INTEGER_TYPE_P(value)) {
        W("\"><v>");
        write_int(buf, NUM2LONG(value));
        W("</v></c>");
    } else if (RB_FLOAT_TYPE_P(value)) {
        W("\"><v>");
        /* Use Ruby's to_s for float to match Ruby path output exactly */
        VALUE fs = rb_funcall(value, rb_intern("to_s"), 0);
        dynbuf_append(buf, RSTRING_PTR(fs), (size_t)RSTRING_LEN(fs));
        W("</v></c>");
    } else {
        /* String, Date, Time — call to_s */
        VALUE s;
        if (rb_respond_to(value, rb_intern("iso8601"))) {
            s = rb_funcall(value, rb_intern("iso8601"), 0);
        } else {
            s = rb_funcall(value, rb_intern("to_s"), 0);
        }
        W("\" t=\"inlineStr\"><is><t>");
        write_escaped(buf, RSTRING_PTR(s), RSTRING_LEN(s));
        W("</t></is></c>");
    }

    #undef W
}

/*
 * Rbxl::Native.generate_sheet(rows) → XML string
 *
 * rows: Array of Arrays, each inner array is a row of cell values
 */
static VALUE rb_native_generate(VALUE self, VALUE rows)
{
    (void)self;
    Check_Type(rows, T_ARRAY);

    VALUE mRbxl = rb_const_get(rb_cObject, rb_intern("Rbxl"));
    VALUE cWriteOnlyCell = rb_const_get(mRbxl, rb_intern("WriteOnlyCell"));

    long num_rows = RARRAY_LEN(rows);

    /* Find max columns for dimension ref */
    int max_cols = 1;
    for (long i = 0; i < num_rows; i++) {
        VALUE row = rb_ary_entry(rows, i);
        int len = (int)RARRAY_LEN(row);
        if (len > max_cols) max_cols = len;
    }

    dynbuf buf;
    dynbuf_init(&buf);

    #define W(s) dynbuf_append(&buf, s, sizeof(s) - 1)

    W("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "  <dimension ref=\"A1:");
    write_column_name(&buf, max_cols);
    write_int(&buf, num_rows);
    W("\"/>\n  <sheetData>");

    for (long i = 0; i < num_rows; i++) {
        VALUE row = rb_ary_entry(rows, i);
        Check_Type(row, T_ARRAY);
        long row_num = i + 1;
        long ncols = RARRAY_LEN(row);

        W("<row r=\"");
        write_int(&buf, row_num);
        W("\">");

        for (long j = 0; j < ncols; j++) {
            write_cell(&buf, (int)(j + 1), (int)row_num, rb_ary_entry(row, j), cWriteOnlyCell);
        }

        W("</row>");
    }

    W("</sheetData>\n</worksheet>");

    #undef W

    VALUE result = make_utf8_str(buf.data, (long)buf.len);
    dynbuf_free(&buf);
    return result;
}

/* ------------------------------------------------------------------ */
/* Init                                                                */
/* ------------------------------------------------------------------ */

void Init_rbxl_native(void)
{
    enc_utf8 = rb_utf8_encoding();
    VALUE mRbxl = rb_define_module("Rbxl");
    VALUE mNative = rb_define_module_under(mRbxl, "Native");
    rb_define_module_function(mNative, "parse_sheet", rb_native_parse, 2);
    rb_define_module_function(mNative, "parse_sheet_full", rb_native_parse_full, 2);
    rb_define_module_function(mNative, "parse_sheet_io", rb_native_parse_io, 3);
    rb_define_module_function(mNative, "parse_sheet_full_io", rb_native_parse_full_io, 3);
    rb_define_module_function(mNative, "generate_sheet", rb_native_generate, 1);
}
