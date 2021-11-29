<template>
  <ion-page>
    <ion-header :translucent="true">
      <ion-toolbar>
        <ion-title>Vue Graph</ion-title>
      </ion-toolbar>
    </ion-header>

    <ion-content :fullscreen="true">
      <div id="container">
        <div v-if="!user">
          <ion-button>Sign in</ion-button>
        </div>
        <div v-if="user">
          <p>Hello, {{ user.name }}</p>
          <ul>
            <li v-for="item in messages" :key="item.id">
              {{ item.subject }}
            </li>
          </ul>
          <ion-button>Sign out</ion-button>
        </div>
      </div>
    </ion-content>
  </ion-page>
</template>

<script>
import {
  IonContent,
  IonHeader,
  IonPage,
  IonTitle,
  IonToolbar,
} from "@ionic/vue";
import { defineComponent } from "vue";

export default defineComponent({
  name: "Home",
  components: {
    IonContent,
    IonHeader,
    IonPage,
    IonTitle,
    IonToolbar,
  },

  data: function () {
    return {
      user: {},
      messages: [
        { subject: "Learn JavaScript" },
        { subject: "Learn Vue.js" },
        { subject: "Build Something Awesome" },
      ],
    };
  },
});
</script>

<style scoped>
#container {
  text-align: center;

  position: absolute;
  left: 0;
  right: 0;
  top: 50%;
  transform: translateY(-50%);
}

#container strong {
  font-size: 20px;
  line-height: 26px;
}

#container p {
  font-size: 16px;
  line-height: 22px;

  color: #8c8c8c;

  margin: 0;
}

#container a {
  text-decoration: none;
}
</style>