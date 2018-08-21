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

/**
 * Excel formula.
 */
class Formula {

    /**
     * Formula expression.
     */
    private final String expression;

    /**
     * Constructor.
     *
     * @param expression Formula expression.
     */
    Formula(String expression) {
        this.expression = expression;
    }

    /**
     * Get formula expression.
     *
     * @return Expression.
     */
    public String getExpression() {
        return expression;
    }

}
