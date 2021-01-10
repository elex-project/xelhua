/*
 * Apache License
 * Version 2.0, January 2004
 * http://www.apache.org/licenses/
 *
 * Copyright (c) 2021, Elex
 * All rights reserved.
 */

plugins {
    java
    `java-library`
    `maven-publish`
    id("com.github.ben-manes.versions") version "0.36.0"
}

group = "com.elex-project"
version = "1.1.0"
description = "Base classes for manipulating Excel file format."

repositories {
    mavenCentral()
    mavenLocal()
}

java {
    withSourcesJar()
    withJavadocJar()
    sourceCompatibility = org.gradle.api.JavaVersion.VERSION_1_8
    targetCompatibility = org.gradle.api.JavaVersion.VERSION_1_8
}

configurations {
    compileOnly {
        //extendsFrom(annotationProcessor.get())
    }
    testCompileOnly {
        //extendsFrom(testAnnotationProcessor.get())
    }
}

tasks.jar {
    manifest {
        attributes(mapOf(
                "Implementation-Title" to project.name,
                "Implementation-Version" to project.version,
                "Implementation-Vendor" to "ELEX co.,pte.",
                "Automatic-Module-Name" to "com.elex_project.dwarf"
        ))
    }
}

tasks.compileJava {
    options.encoding = "UTF-8"
}

tasks.compileTestJava {
    options.encoding = "UTF-8"
}

tasks.test {
    useJUnitPlatform()
}

tasks.javadoc {
    if (JavaVersion.current().isJava9Compatible) {
        (options as StandardJavadocDocletOptions).addBooleanOption("html5", true)
    }
    (options as StandardJavadocDocletOptions).encoding = "UTF-8"
    (options as StandardJavadocDocletOptions).charSet = "UTF-8"
    (options as StandardJavadocDocletOptions).docEncoding = "UTF-8"

}

publishing {
    publications {
        create<MavenPublication>("mavenJava") {
            from(components["java"])
            pom {
                name.set(project.name)
                description.set(project.description)
                url.set("https://github.com/elex-project/xelhua")
                licenses {
                    license {
                        name.set("Apache-2.0 License")
                        url.set("https://github.com/elex-project/xelhua/blob/main/LICENSE")
                    }
                }
                developers {
                    developer {
                        id.set("elex-project")
                        name.set("Elex")
                        email.set("developer@elex-project.com")
                    }
                }
                scm {
                    connection.set("scm:git:https://github.com/elex-project/xelhua.git")
                    developerConnection.set("scm:git:https://github.com/elex-project/xelhua.git")
                    url.set("https://github.com/elex-project/xelhua")
                }
            }
        }
    }

    repositories {
        maven {
            name = "mavenGithub"
            url = uri("https://maven.pkg.github.com/elex-project/xelhua")
            credentials {
                username = project.findProperty("github.username") as String
                password = project.findProperty("github.token") as String
            }
        }
    }
}
dependencies {
    implementation("org.jetbrains:annotations:20.1.0")

    api("org.apache.poi:poi:4.1.2")
    api("org.apache.poi:poi-ooxml:4.1.2")

    testImplementation("org.junit.jupiter:junit-jupiter:5.7.0")
    testRuntimeOnly("org.junit.jupiter:junit-jupiter-engine:5.7.0")

}
