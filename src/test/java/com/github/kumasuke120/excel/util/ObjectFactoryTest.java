package com.github.kumasuke120.excel.util;

import org.junit.jupiter.api.Test;

import java.math.BigDecimal;

import static org.junit.jupiter.api.Assertions.*;


class ObjectFactoryTest {

    @Test
    void newInstanceHandle() {
        final ObjectFactory<A> aHandleFactory = ObjectFactory.buildFactory(A.class, true);
        assertNotNull(aHandleFactory);
        final A a = aHandleFactory.newInstance();
        assertNotNull(a);

        final ObjectFactory<B> bHandleFactory = ObjectFactory.buildFactory(B.class, true);
        assertNotNull(bHandleFactory);
        final B b = bHandleFactory.newInstance();
        assertNotNull(b);

        final ObjectFactory<E> eHandleFactory = ObjectFactory.buildFactory(E.class, true);
        assertNotNull(eHandleFactory);
        assertThrows(ObjectCreationException.class, eHandleFactory::newInstance);

        assertThrows(ObjectCreationException.class, () -> ObjectFactory.buildFactory(P.class, true));

        assertThrows(ObjectCreationException.class, () -> ObjectFactory.buildFactory(N.class, true));
    }

    @Test
    void newInstanceReflect() {
        final ObjectFactory<A> aHandleFactory = ObjectFactory.buildFactory(A.class, false);
        assertNotNull(aHandleFactory);
        final A a = aHandleFactory.newInstance();
        assertNotNull(a);

        final ObjectFactory<B> bHandleFactory = ObjectFactory.buildFactory(B.class, false);
        assertNotNull(bHandleFactory);
        final B b = bHandleFactory.newInstance();
        assertNotNull(b);

        final ObjectFactory<E> eHandleFactory = ObjectFactory.buildFactory(E.class, false);
        assertNotNull(eHandleFactory);
        assertThrows(ObjectCreationException.class, eHandleFactory::newInstance);

        assertThrows(ObjectCreationException.class, () -> ObjectFactory.buildFactory(P.class, false));

        assertThrows(ObjectCreationException.class, () -> ObjectFactory.buildFactory(N.class, false));
    }


    @Test
    void newInstanceSpeedTest() {
        // heat up
        newInstanceSpeedTest(A.class, 500_000, true, true);
        newInstanceSpeedTest(A.class, 500_000, false, true);

        newInstanceSpeedTest(B.class, 500_000, true, true);
        newInstanceSpeedTest(B.class, 500_000, false, true);

        // actual test
        newInstanceSpeedTest(A.class, 100_000, true);
        newInstanceSpeedTest(B.class, 100_000, true);

        newInstanceSpeedTest(A.class, 100_000, false);
        newInstanceSpeedTest(B.class, 100_000, false);

        newInstanceSpeedTest(A.class, 5_000_000, true);
        newInstanceSpeedTest(B.class, 5_000_000, true);

        newInstanceSpeedTest(A.class, 5_000_000, false);
        newInstanceSpeedTest(B.class, 5_000_000, false);

        newInstanceSpeedTest(A.class, 50_000_000, true);
        newInstanceSpeedTest(B.class, 50_000_000, true);

        newInstanceSpeedTest(A.class, 50_000_000, false);
        newInstanceSpeedTest(B.class, 50_000_000, false);
    }

    private <T> void newInstanceSpeedTest(Class<T> clazz, int testCount, boolean useHandle) {
        newInstanceSpeedTest(clazz, testCount, useHandle, false);
    }

    private <T> void newInstanceSpeedTest(Class<T> clazz, int testCount, boolean useHandle, boolean heatUp) {
        final ObjectFactory<T> factory = ObjectFactory.buildFactory(clazz, useHandle);
        assertNotNull(factory);

        long start = System.nanoTime();
        for (int i = 0; i < testCount; i++) {
            final T t = factory.newInstance();
            assertNotNull(t);
        }

        if (!heatUp) {
            final double elapsed = (System.nanoTime() - start) / 1_000_000d;
            System.out.printf("[%s][%s] %s times new instance cost %.3f ms\n", clazz.getSimpleName(),
                    useHandle ? "handle" : "reflect", abbrNum(testCount), elapsed);
        }

        System.gc();
    }

    private static String abbrNum(int number) {
        if (number < 1000) {
            return String.valueOf(number);
        } else if (number < 1000000) {
            return (number / 1000) + "k";
        } else if (number < 1000000000) {
            return (number / 1000000) + "M";
        } else {
            return (number / 1000000000) + "B";
        }
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