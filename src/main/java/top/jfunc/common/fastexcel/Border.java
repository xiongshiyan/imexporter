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
import java.util.Objects;

/**
 * Border attributes.
 */
class Border {

    /**
     * Default border attributes: no borders.
     */
    static final Border NONE = new Border(BorderElement.NONE, BorderElement.NONE, BorderElement.NONE, BorderElement.NONE, BorderElement.NONE);

    /**
     * Left border element.
     */
    private final BorderElement left;
    /**
     * Right border element.
     */
    private final BorderElement right;
    /**
     * Top border element.
     */
    private final BorderElement top;
    /**
     * Bottom border element.
     */
    private final BorderElement bottom;
    /**
     * Diagonal border element.
     */
    private final BorderElement diagonal;

    /**
     * Constructor.
     *
     * @param left Border element for left side.
     * @param right Border element for right side.
     * @param top Border element for top side.
     * @param bottom Border element for bottom side.
     * @param diagonal Border element for diagonal side.
     */
    Border(BorderElement left, BorderElement right, BorderElement top, BorderElement bottom, BorderElement diagonal) {
        this.left = left;
        this.right = right;
        this.top = top;
        this.bottom = bottom;
        this.diagonal = diagonal;
    }

    /**
     * Create a border where all sides have the same style and color.
     *
     * @param style Border style. Possible values are defined <a href="https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.borderstylevalues(v=office.14).aspx">here</a>.
     * @param color RGB border color.
     * @return A new border object.
     */
    static Border fromStyleAndColor(String style, String color) {
        BorderElement element = new BorderElement(style, color);
        return new Border(element, element, element, element, BorderElement.NONE);
    }

    @Override
    public int hashCode() {
        int hash = 3;
        hash = 17 * hash + Objects.hashCode(this.left);
        hash = 17 * hash + Objects.hashCode(this.right);
        hash = 17 * hash + Objects.hashCode(this.top);
        hash = 17 * hash + Objects.hashCode(this.bottom);
        hash = 17 * hash + Objects.hashCode(this.diagonal);
        return hash;
    }

    @Override
    public boolean equals(Object obj) {
        if (obj == null) {
            return false;
        }
        if (getClass() != obj.getClass()) {
            return false;
        }
        final Border other = (Border) obj;
        if (!Objects.equals(this.left, other.left)) {
            return false;
        }
        if (!Objects.equals(this.right, other.right)) {
            return false;
        }
        if (!Objects.equals(this.top, other.top)) {
            return false;
        }
        if (!Objects.equals(this.bottom, other.bottom)) {
            return false;
        }
        if (!Objects.equals(this.diagonal, other.diagonal)) {
            return false;
        }
        return true;
    }

    /**
     * Write this border definition as an XML element.
     *
     * @param w Output writer.
     * @throws IOException If an I/O error occurs.
     */
    void write(Writer w) throws IOException {
        w.append("<border>");
        left.write("left", w);
        right.write("right", w);
        top.write("top", w);
        bottom.write("bottom", w);
        diagonal.write("diagonal", w);
        w.append("</border>");
    }

}
