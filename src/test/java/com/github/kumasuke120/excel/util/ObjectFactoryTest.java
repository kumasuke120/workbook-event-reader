package com.github.kumasuke120.excel.util;

import org.junit.jupiter.api.Test;

import java.math.BigDecimal;

import static org.junit.jupiter.api.Assertions.*;


class ObjectFactoryTest {

    @Test
    void newInstance() {
        final ObjectFactory<A> aHandleFactory = ObjectFactory.buildFactory(A.class);
        assertNotNull(aHandleFactory);
        final A a = aHandleFactory.newInstance();
        assertNotNull(a);

        final ObjectFactory<B> bHandleFactory = ObjectFactory.buildFactory(B.class);
        assertNotNull(bHandleFactory);
        final B b = bHandleFactory.newInstance();
        assertNotNull(b);

        final ObjectFactory<E> eHandleFactory = ObjectFactory.buildFactory(E.class);
        assertNotNull(eHandleFactory);
        assertThrows(ObjectCreationException.class, eHandleFactory::newInstance);

        final ObjectFactory<EA> eaHandleFactory = ObjectFactory.buildFactory(EA.class);
        assertNotNull(eaHandleFactory);
        assertThrows(AssertionError.class, eaHandleFactory::newInstance);

        assertThrows(ObjectCreationException.class, () -> ObjectFactory.buildFactory(P.class));

        assertThrows(ObjectCreationException.class, () -> ObjectFactory.buildFactory(N.class));
    }

    public static class A {
    }

    @SuppressWarnings("unused")
    public static class B {

        private int a;
        private double b;
        private Object c;
        private String d;
        private BigDecimal e;

    }

    public static class E {
        public E() {
            throw new UnsupportedOperationException();
        }
    }

    public static class EA {
        public EA() {
            throw new AssertionError();
        }
    }

    public static class P {
        private P() {
        }
    }

    @SuppressWarnings("unused")
    public static class N {
        final int n;

        public N(int n) {
            this.n = n;
        }
    }

}