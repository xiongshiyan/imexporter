/*
 * Copyright 2016 Dhatim.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package top.jfunc.common.fastexcel;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;

/**
 * Writer for XML files.
 */
class Writer {

    /**
     * Target output stream.
     */
    private final OutputStream os;
    /**
     * Char buffer.
     */
    private final StringBuffer sb;

    /**
     * Constructor.
     *
     * @param os Output stream.
     */
    Writer(OutputStream os) {
        this.os = os;
        this.sb = new StringBuffer(512 * 1024);
    }

    /**
     * Append a string without escaping.
     *
     * @param s String.
     * @return This writer.
     * @throws IOException If an I/O error occurs.
     */
    Writer append(String s) throws IOException {
        return append(s, false);
    }

    /**
     * Append a string with XML escaping.
     *
     * @param s String.
     * @return This writer.
     * @throws IOException If an I/O error occurs.
     */
    Writer appendEscaped(String s) throws IOException {
        return append(s, true);
    }

    /**
     * Append a char with escaping.
     *
     * @param c Character.
     * @throws IOException If an I/O error occurs.
     */
    private void escape(char c) throws IOException {
        switch (c) {
            case '<':
                sb.append("&lt;");
                break;
            case '>':
                sb.append("&gt;");
                break;
            case '&':
                sb.append("&amp;");
                break;
            case '\'':
                sb.append("&apos;");
                break;
            case '"':
                sb.append("&quot;");
                break;
            default:
                if (c > 0x7e) {
                    sb.append("&#x").append(Integer.toHexString(c)).append(';');
                } else {
                    sb.append(c);
                }
                break;
        }
    }

    /**
     * Append a string with or without escaping.
     *
     * @param s String.
     * @param escape Is the string escaped?
     * @return This writer.
     * @throws IOException If an I/O error occurs.
     */
    private Writer append(String s, boolean escape) throws IOException {
        if (escape) {
            for (int i = 0; i < s.length(); ++i) {
                escape(s.charAt(i));
            }
        } else {
            sb.append(s);
        }
        check();
        return this;
    }

    /**
     * Check if the buffer gets full. In this case, flush bytes to the output
     * stream.
     *
     * @throws IOException If an I/O error occurs.
     */
    private void check() throws IOException {
        if (sb.capacity() - sb.length() < 1024) {
            flush();
        }
    }

    /**
     * Append a char without escaping.
     *
     * @param c Character.
     * @return This writer.
     * @throws IOException If an I/O error occurs.
     */
    Writer append(char c) throws IOException {
        sb.append(c);
        check();
        return this;
    }

    /**
     * Append a boolean.
     *
     * @param b Boolean.
     * @return This writer.
     * @throws IOException If an I/O error occurs.
     */
    Writer append(boolean b) throws IOException {
        sb.append(b);
        check();
        return this;
    }

    /**
     * Append an integer.
     *
     * @param n Integer.
     * @return This writer.
     * @throws IOException If an I/O error occurs.
     */
    Writer append(int n) throws IOException {
        sb.append(n);
        check();
        return this;
    }

    /**
     * Append a long.
     *
     * @param n Long.
     * @return This writer.
     * @throws IOException If an I/O error occurs.
     */
    Writer append(long n) throws IOException {
        sb.append(n);
        check();
        return this;
    }

    /**
     * Append a double.
     *
     * @param n Double.
     * @return This writer.
     * @throws IOException If an I/O error occurs.
     */
    Writer append(double n) throws IOException {
        sb.append(n);
        check();
        return this;
    }

    /**
     * Flush this writer.
     *
     * @throws IOException If an I/O error occurs.
     */
    void flush() throws IOException {
        os.write(sb.toString().getBytes(StandardCharsets.UTF_8));
        sb.setLength(0);
    }
}
